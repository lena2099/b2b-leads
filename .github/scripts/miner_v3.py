#!/usr/bin/env python3
"""
B2B Lead Miner v3 — Rhythmic deep-research pipeline.
Hourly trigger:
  - Every run: deep-research 2 incomplete buyers
  - Every 3rd run: + generate 5 new buyers  
  - Every 6th run: + rebuild dashboard/Excel/Obsidian
"""
import json, os, hashlib, time
from datetime import datetime, timezone
from pathlib import Path
from urllib.request import Request, urlopen

API_KEY = os.environ["DEEPSEEK_API_KEY"]
DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
STATE_FILE = DATA_DIR / "miner_state.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

INDUSTRIES = {
    "robotics":            {"cn":"机器人&自动化设备","emoji":"🤖","regions":["美国","德国","日本","韩国","加拿大","英国"],"types":["汽车制造商","电子厂","3PL物流","系统集成商","食品加工厂","钣金焊接厂","制药厂"]},
    "energy_storage":      {"cn":"储能系统&电池","emoji":"🔋","regions":["美国","德国","英国","澳大利亚","荷兰","西班牙"],"types":["电力公司","光伏安装商","可再生能源开发商","EV充电运营商","商业地产商"]},
    "solar_pv":            {"cn":"光伏组件&逆变器","emoji":"☀️","regions":["美国","德国","巴西","印度","澳大利亚","荷兰"],"types":["光伏安装商","EPC工程商","电力公司","光伏分销商"]},
    "medical_device":      {"cn":"医疗器械&设备","emoji":"🏥","regions":["美国","德国","巴西","印尼","尼日利亚","墨西哥"],"types":["医院集团","医疗器械分销商","GPO采购","诊所连锁"]},
    "laser_equipment":     {"cn":"激光设备&精密加工","emoji":"🔩","regions":["美国","德国","日本","意大利","墨西哥","土耳其"],"types":["钣金加工厂","汽车零部件厂","电子代工厂","设备分销商"]},
    "ev_charger":          {"cn":"新能源汽车&充电桩","emoji":"🚗","regions":["德国","挪威","荷兰","泰国","巴西","阿联酋"],"types":["汽车经销商","车队运营商","租车公司","充电网络运营商"]},
    "construction_machinery":{"cn":"工程机械&农业机械","emoji":"🏗️","regions":["印尼","印度","巴西","沙特","阿联酋","尼日利亚"],"types":["建筑公司","设备租赁公司","矿业公司","大型农场","政府招标"]},
    "telecom_iot":         {"cn":"通信设备&物联网","emoji":"📡","regions":["巴西","印尼","尼日利亚","沙特","马来西亚","泰国"],"types":["电信运营商","ISP","系统集成商","智慧城市承包商","安防公司"]},
}

DEEP_FIELDS = ["founded","headquarters","employees","annual_revenue","business_description",
               "procurement_needs","partner_profile","key_markets","china_presence",
               "competitors","market_share","recent_news"]

# ══════════════════════════════════════════════════════════
# LLM
# ══════════════════════════════════════════════════════════
def ask(message, max_tokens=3500, temp=0.4):
    body = json.dumps({
        "model":"deepseek-chat","messages":[{"role":"user","content":message}],
        "max_tokens":max_tokens,"temperature":temp
    }).encode()
    req = Request("https://api.deepseek.com/chat/completions", data=body,
                  headers={"Authorization":f"Bearer {API_KEY}","Content-Type":"application/json"})
    resp = json.loads(urlopen(req, timeout=120).read())
    return resp["choices"][0]["message"]["content"].strip()

# ══════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"total_runs":0,"total_buyers":0,"seen":[],"last_run":None,"rotation":list(INDUSTRIES.keys()),"rot_idx":0}

def save_state(s):
    s["total_runs"] += 1
    s["last_run"] = datetime.now(timezone.utc).isoformat()
    STATE_FILE.write_text(json.dumps(s, indent=2, ensure_ascii=False))

# ══════════════════════════════════════════════════════════
# STEP 1: DEEP RESEARCH — fill in missing profiles
# ══════════════════════════════════════════════════════════
def deep_research_2(ik, cfg):
    """Find 2 buyers with incomplete profiles in this industry and fill them."""
    fp = OUTPUT_DIR / f"{ik}_leads.json"
    if not fp.exists():
        return 0
    
    data = json.loads(fp.read_text())
    leads = data["leads"]
    
    # Find incomplete ones
    incomplete = []
    for l in leads:
        missing = [f for f in DEEP_FIELDS if not l.get(f) or str(l.get(f)).strip() in ("待补充","","[]","{}")]
        if missing:
            incomplete.append((l, missing))
    
    if not incomplete:
        print(f"  ✅ all {len(leads)} profiles complete, skip")
        return 0
    
    # Take 2 most relevant incomplete ones
    targets = incomplete[:2]
    print(f"  🔬 Deep-researching {len(targets)} of {len(incomplete)} incomplete profiles...")
    
    for l, missing_fields in targets:
        name = l["company_name"]
        existing_info = json.dumps({k: l.get(k) for k in DEEP_FIELDS if l.get(k) and str(l.get(k)).strip() not in ('待补充','','[]','{}')}, ensure_ascii=False)
        json_template = '{' + ', '.join(f'"{f}": "..."' for f in missing_fields) + '}'
        
        prompt = f"""请为"{name}"撰写一份专业的企业档案（中文，简洁要点形式）。根据你的知识库中关于这家公司的信息。

需要覆盖的维度：{', '.join(missing_fields)}

当前已知的部分信息（如有）：{existing_info}

行业背景：{cfg['cn']}的海外买家。
要求：
- 所有信息基于真实数据，不确定的标注"数据待确认"
- 每个字段1-3句话，不要冗长
- 返回JSON格式：{json_template}

只返回JSON。"""
        
        try:
            result = ask(prompt, max_tokens=2000, temp=0.3)
            result = result.replace("```json","").replace("```","").strip()
            enriched = json.loads(result)
            for k, v in enriched.items():
                if v and k in DEEP_FIELDS and v != "待补充":
                    l[k] = v
            print(f"    ✅ {name}: filled {len(enriched)} fields")
        except Exception as e:
            print(f"    ⚠️ {name}: {str(e)[:80]}")
            # Mark as attempted to avoid infinite retry
            for f in missing_fields:
                if not l.get(f): l[f] = "数据待确认"
        time.sleep(0.5)
    
    # Save back
    with open(fp, "w") as f:
        json.dump({"industry":cfg["cn"],"total":len(leads),"updated":datetime.now(timezone.utc).isoformat(),"leads":leads}, f, indent=2, ensure_ascii=False)
    
    return len(targets)

# ══════════════════════════════════════════════════════════
# STEP 2: GENERATE NEW BUYERS (every 3rd run)
# ══════════════════════════════════════════════════════════
def generate_new_buyers(state):
    """Generate 5 new buyers with basic profiles from next industry."""
    ik = state["rotation"][state["rot_idx"]]
    cfg = INDUSTRIES[ik]
    state["rot_idx"] = (state["rot_idx"] + 1) % len(INDUSTRIES)
    
    print(f"  🧠 Generating 5 new buyers for {cfg['emoji']} {cfg['cn']}...")
    
    prompt = f"""列出5家真实的海外{cfg['cn']}买家。必须在2026年存续经营。
国家: {', '.join(cfg['regions'][:5])}
类型: {'; '.join(cfg['types'][:4])}

每家公司写完整档案（中文），包含：
company_name, website, founded, headquarters, employees, annual_revenue,
business_description, procurement_needs(需要从中国采购什么),
partner_profile(要什么供应商), key_markets(主要市场),
china_presence(在华分支,没有写"暂无"), competitors, market_share,
recent_news(2025-2026最新动态), buyer_type, country, size, relevance(45-95), why

只返回JSON数组。字段尽量精简。"""
    
    try:
        result = ask(prompt, max_tokens=4000)
        result = result.replace("```json","").replace("```","").strip()
        buyers = json.loads(result)
    except Exception as e:
        print(f"    ❌ Generation failed: {e}")
        return 0
    
    # Dedup
    seen = set(state.get("seen",[]))
    new = []
    for b in buyers:
        h = hashlib.md5(b["company_name"].lower().strip().encode()).hexdigest()
        if h not in seen:
            seen.add(h)
            b["id"] = h[:12]; b["industry"] = ik; b["generated_at"] = datetime.now(timezone.utc).isoformat()
            new.append(b)
    state["seen"] = list(seen)
    
    if not new:
        print("    No new buyers after dedup")
        return 0
    
    fpath = OUTPUT_DIR / f"{ik}_leads.json"
    existing = []
    if fpath.exists():
        try: existing = json.loads(fpath.read_text()).get("leads",[])
        except: pass
    existing.extend(new)
    with open(fpath, "w") as f:
        json.dump({"industry":cfg["cn"],"total":len(existing),"updated":datetime.now(timezone.utc).isoformat(),"leads":existing}, f, indent=2, ensure_ascii=False)
    
    state["total_buyers"] += len(new)
    print(f"    ✅ Added {len(new)} new buyers (total: {len(existing)})")
    return len(new)

# ══════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════
def main():
    run_id = int(time.strftime("%H"))
    print(f"\n{'='*60}")
    print(f"  🌍 B2B Lead Miner v3 — Run #{run_id}")
    print(f"  {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
    
    state = load_state()
    researched = 0; added = 0
    
    # ── STEP 1: Deep research (every run) ──
    ik = state["rotation"][state["rot_idx"] % len(INDUSTRIES)]
    # Rotate through industries for deep research
    state["rot_idx"] = (state["rot_idx"] + 1) % len(INDUSTRIES)
    cfg = INDUSTRIES[ik]
    print(f"\n  🔬 Industry: {cfg['emoji']} {cfg['cn']}")
    researched = deep_research_2(ik, cfg)
    
    # ── STEP 2: New buyers (every 3rd run) ──
    if run_id % 3 == 0:
        print(f"\n  🧠 Generation cycle (run#{run_id} % 3 = 0)")
        added = generate_new_buyers(state)
    
    save_state(state)
    
    # Print stats
    total_complete = 0; total_all = 0
    for ik2 in INDUSTRIES:
        fp = OUTPUT_DIR / f"{ik2}_leads.json"
        if fp.exists():
            leads = json.loads(fp.read_text())["leads"]
            c = sum(1 for l in leads if all(l.get(f) and str(l.get(f)).strip() not in ("待补充","","[]","{}") for f in DEEP_FIELDS))
            total_complete += c; total_all += len(leads)
    
    print(f"\n  📊 {total_complete}/{total_all} profiles complete ({total_complete*100//total_all}%)")
    print(f"  🔬 Researched: {researched} | 🧠 Added: {added}")
    print(f"  ⏭️  Next run in ~1 hour\n")

if __name__ == "__main__":
    main()
