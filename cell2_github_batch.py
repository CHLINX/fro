#!/usr/bin/env python3
# FRO Cell2 GitHub Actions batch runner
# Based on Cell2 v8.6 NLSC browser-session + qryTileMapIndex logic.
import asyncio, csv, json, os, re, sys, time
from pathlib import Path
from datetime import datetime, timezone
from playwright.async_api import async_playwright

NLSC_URL = "https://maps.nlsc.gov.tw/T09/mobilemap.action"
ENDPOINT = "https://landmaps.nlsc.gov.tw/S_Maps/qryTileMapIndex"
OUT = Path("output")
OUT.mkdir(exist_ok=True)
RAW = OUT / "raw"
RAW.mkdir(exist_ok=True)
CHECKPOINT = OUT / "checkpoint.jsonl"
MIN_INTERVAL = 0.8
MAX_RETRIES = 5
COOLDOWN_UNTIL = 0.0

def clean(v):
    return re.sub(r"\s+", "", str(v or "").strip())

def norm_land(v):
    s = clean(v)
    if re.fullmatch(r"\d+\.0", s): s = s[:-2]
    if "-" in s or "－" in s:
        p = s.replace("－","-").split("-")
        if len(p)==2 and all(x.isdigit() for x in p):
            return p[0].zfill(4)+p[1].zfill(4)
        return s
    if s.isdigit():
        if len(s)<=4: return s.zfill(4)+"0000"
        if len(s)<8: return s.zfill(8)
    return s

def canon(s):
    return clean(s).replace("小段","").replace("段","")

def section_match(req, actual):
    r,a=clean(req),clean(actual)
    if not r or not a: return False
    if r==a or canon(r)==canon(a): return True
    variants=lambda x:{x, x[:-2] if x.endswith("小段") else x,
                       x[:-1] if x.endswith("段") else x,
                       x+"段" if not x.endswith("段") else x,
                       x+"小段" if not x.endswith("小段") else x}
    if variants(r)&variants(a): return True
    if "?" in r:
        return re.fullmatch(re.escape(r).replace(r"\?", "."),a) is not None
    return False

def area_match(req, actual):
    r,a=clean(req),clean(actual)
    strip=lambda x: re.sub(r"(鄉|鎮|市|區)$","",x)
    return r==a or strip(r)==strip(a)

def parse_jsonp(body):
    s=(body or "").strip()
    if s.startswith("[") or s.startswith("{"): return json.loads(s)
    m=re.match(r"^[^(]+\((.*)\)\s*;?\s*$",s,re.S)
    if not m: raise ValueError("invalid JSON/JSONP")
    return json.loads(m.group(1))

async def options(page, selector):
    return await page.eval_on_selector_all(selector,
        "els=>els.map(e=>({text:(e.textContent||'').trim(),value:e.value||''}))")

async def wait_options(page, selector, predicate, timeout=40):
    end=time.monotonic()+timeout
    while time.monotonic()<end:
        try:
            opts=await options(page,selector)
            found=[x for x in opts if x["value"] and x["text"] and predicate(x)]
            if found: return found
        except Exception: pass
        await page.wait_for_timeout(500)
    raise TimeoutError(f"option wait timeout: {selector}")

class NLSC:
    def __init__(self,page):
        self.page=page
        self.area_cache={}
        self.last_req=0.0

    async def section_info(self, city, area, section):
        key=(city,area)
        if key not in self.area_cache:
            await self.page.select_option("#land_city",label=city,force=True)
            aopts=await wait_options(self.page,"#land_area_office option",
                                     lambda x:area_match(area,x["text"]))
            if len(aopts)!=1:
                raise RuntimeError(f"鄉鎮候選不唯一：{city}/{area}: {[x['text'] for x in aopts]}")
            a=aopts[0]
            await self.page.evaluate("""({value})=>{
              const e=document.querySelector('#land_area_office');
              e.value=value;e.dispatchEvent(new Event('change',{bubbles:true}));
            }""",{"value":a["value"]})
            # Wait for section options to load for this area.
            await wait_options(self.page,"#land_section option",
                               lambda x: bool(x["value"]),timeout=45)
            self.area_cache[key]={"area_text":a["text"],"area_value":a["value"],
                                  "sections":await options(self.page,"#land_section option")}
        cat=self.area_cache[key]
        matches=[x for x in cat["sections"] if x["value"] and x["text"] and section_match(section,x["text"])]
        if len(matches)!=1:
            raise RuntimeError(f"地段候選非唯一/不存在：{city}/{area}/{section}: {[x['text'] for x in matches]}")
        opt=matches[0]
        # If switching town/area, restore the cached area DOM first.
        current_city=await self.page.input_value("#land_city")
        if current_city!=city:
            await self.page.select_option("#land_city",label=city,force=True)
        current_area=await self.page.input_value("#land_area_office")
        if current_area!=cat["area_value"]:
            await self.page.evaluate("""({value})=>{
              const e=document.querySelector('#land_area_office');
              e.value=value;e.dispatchEvent(new Event('change',{bubbles:true}));
            }""",{"value":cat["area_value"]})
            await wait_options(self.page,"#land_section option",lambda x: x["value"]==opt["value"],timeout=45)
        await self.page.evaluate("""({value})=>{
          const e=document.querySelector('#land_section');
          e.value=value;e.dispatchEvent(new Event('change',{bubbles:true}));
        }""",{"value":opt["value"]})
        code=""
        for _ in range(40):
            try:
                code=clean(await self.page.input_value("#land_sectioncode"))
                if code: break
            except Exception: pass
            await self.page.wait_for_timeout(250)
        if not code: raise RuntimeError(f"未取得 sectioncode：{section}")
        value=opt["value"]
        office=value.split("_",1)[0] if "_" in value else ""
        sect=code or (value.split("_",1)[1] if "_" in value else value)
        if not office or not sect:
            raise RuntimeError(f"office/sect 無效：value={value}, code={code}")
        return cat["area_text"],opt["text"],office,sect

    async def query(self,row,index):
        global COOLDOWN_UNTIL
        city=clean(row.get("縣市")); area=clean(row.get("鄉鎮市區"))
        section=clean(row.get("地段")); land=clean(row.get("地號"))
        normalized=norm_land(land)
        area_text,section_text,office,sect=await self.section_info(city,area,section)
        last_error=""
        for attempt in range(1,MAX_RETRIES+1):
            now=time.monotonic()
            if now<COOLDOWN_UNTIL:
                await self.page.wait_for_timeout(int((COOLDOWN_UNTIL-now)*1000))
            gap=MIN_INTERVAL-(time.monotonic()-self.last_req)
            if gap>0: await self.page.wait_for_timeout(int(gap*1000))
            nonce=int(time.time()*1000)
            url=(f"{ENDPOINT}?type=2&flag=2&office={office}&sect={sect}"
                 f"&landno={normalized}&alpah=0.5f&callback=_jqjsp&_{nonce}=")
            self.last_req=time.monotonic()
            try:
                resp=await self.page.context.request.get(url,timeout=60000,
                    headers={"Referer":NLSC_URL,"Accept":"*/*"},
                    fail_on_status_code=False)
                body=await resp.text()
                (RAW/f"{index:03d}_{normalized}_try{attempt}.txt").write_text(
                    f"HTTP {resp.status}\n{resp.url}\n\n{body}",encoding="utf-8")
                if resp.status==429:
                    ra=resp.headers.get("retry-after","")
                    try: delay=min(90,max(10,float(ra)))
                    except Exception: delay=min(90,30*attempt)
                    COOLDOWN_UNTIL=max(COOLDOWN_UNTIL,time.monotonic()+delay)
                    last_error=f"HTTP_429 cooldown={delay}"
                    continue
                if resp.status in (502,503,504):
                    last_error=f"HTTP_{resp.status}"
                    await self.page.wait_for_timeout(min(30000,3000*attempt))
                    continue
                if resp.status!=200:
                    last_error=f"HTTP_{resp.status}: {body[:300]}"
                    continue
                data=parse_jsonp(body)
                rec=data[0][0]
                lon=float(rec["cx"]); lat=float(rec["cy"])
                if not (119.0<=lon<=123.0 and 21.0<=lat<=26.0):
                    raise ValueError(f"座標超出台灣合理範圍：{lon},{lat}")
                return {"input_index":index,"縣市":city,"鄉鎮市區":area_text,
                    "地段":section_text,"地號":land,"查詢地號":normalized,
                    "office":office,"sect":sect,"經度":lon,"緯度":lat,
                    "左下經度":rec.get("lx"),"左下緯度":rec.get("ly"),
                    "右上經度":rec.get("rx"),"右上緯度":rec.get("ry"),
                    "狀態":"OK","HTTP狀態":resp.status,"嘗試次數":attempt,
                    "response_url":resp.url}
            except Exception as e:
                last_error=f"{type(e).__name__}: {e}"
                if attempt<MAX_RETRIES:
                    await self.page.wait_for_timeout(min(20000,1500*attempt))
        return {"input_index":index,"縣市":city,"鄉鎮市區":area,"地段":section,
                "地號":land,"查詢地號":normalized,"狀態":"ERROR",
                "HTTP狀態":"","嘗試次數":MAX_RETRIES,"錯誤":last_error}

async def main():
    source=Path(os.environ.get("PARCELS_FILE","input/parcels.json"))
    if not source.exists():
        raise FileNotFoundError(f"找不到批次輸入檔：{source}")
    rows=json.loads(source.read_text(encoding="utf-8"))
    if not isinstance(rows,list) or not rows: raise ValueError("parcels.json 必須是非空陣列")
    # Remove exact duplicate parcels but preserve original input order and all input indices.
    print(f"批次地號：{len(rows)} 筆")
    results=[]
    async with async_playwright() as p:
        browser=await p.chromium.launch(headless=True,args=[
            "--disable-blink-features=AutomationControlled","--no-sandbox","--disable-dev-shm-usage"])
        context=await browser.new_context(viewport={"width":1440,"height":1000},
            locale="zh-TW",ignore_https_errors=True)
        page=await context.new_page()
        print("開啟 NLSC 並建立瀏覽器工作階段…")
        await page.goto(NLSC_URL,wait_until="domcontentloaded",timeout=90000)
        for sel in ["#land_city","#land_area_office","#land_section","#land_sectioncode"]:
            await page.wait_for_selector(sel,state="attached",timeout=60000)
        await page.wait_for_timeout(1500)
        client=NLSC(page)
        for i,row in enumerate(rows,1):
            if not isinstance(row,dict): continue
            try:
                result=await client.query(row,i)
            except Exception as e:
                result={"input_index":i,**row,"狀態":"ERROR","錯誤":f"{type(e).__name__}: {e}"}
            results.append(result)
            with CHECKPOINT.open("a",encoding="utf-8") as f:
                f.write(json.dumps(result,ensure_ascii=False)+"\n")
            ok=sum(x.get("狀態")=="OK" for x in results)
            print(f"[{i}/{len(rows)}] {row.get('地段')} {row.get('地號')} -> {result.get('狀態')} | 成功 {ok}/{i}")
        await browser.close()
    (OUT/"results.json").write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding="utf-8")
    if results:
        with (OUT/"results.csv").open("w",newline="",encoding="utf-8-sig") as f:
            keys=list(dict.fromkeys(k for x in results for k in x))
            w=csv.DictWriter(f,fieldnames=keys);w.writeheader();w.writerows(results)
    ok=sum(x.get("狀態")=="OK" for x in results)
    print(f"完成：{ok}/{len(results)} 筆取得座標；失敗 {len(results)-ok} 筆")
    if ok==0: sys.exit(2)

if __name__=="__main__":
    asyncio.run(main())
