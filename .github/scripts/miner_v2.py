#!/usr/bin/env python3
"""
B2B Lead Miner v2 — Enhanced Buyer Profiles with 10-Dimension Analysis.
7×24 hourly via GitHub Actions. Each run: picks next industry, generates 10 new buyers.
Outputs: JSON, Excel, bilingual HTML dashboard, Obsidian markdown vault.
"""
import json, os, re, hashlib, time
from datetime import datetime, timezone
from pathlib import Path
from urllib.request import Request, urlopen

API_KEY = os.environ["DEEPSEEK_API_KEY"]
BASE_DIR = Path(".")
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
STATE_FILE = DATA_DIR / "miner_state.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ══════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════
INDUSTRIES = {
    "robotics": {
        "name_cn": "机器人&自动化设备", "name_en": "Robotics & Automation",
        "emoji": "🤖",
        "regions": ["美国","德国","日本","韩国","加拿大","英国"],
        "buyer_types": ["汽车制造商","电子代工厂","3PL物流商","食品加工厂","钣金焊接厂","系统集成商","制药厂"],
        "folder": "机器人行业",
    },
    "energy_storage": {
        "name_cn": "储能系统&电池", "name_en": "Energy Storage & Battery",
        "emoji": "🔋",
        "regions": ["美国","德国","英国","澳大利亚","荷兰","西班牙"],
        "buyer_types": ["电力公司","光伏安装商","可再生能源开发商","EV充电运营商","商业地产商","微电网开发商"],
        "folder": "储能行业",
    },
    "solar_pv": {
        "name_cn": "光伏组件&逆变器", "name_en": "Solar PV & Inverters",
        "emoji": "☀️",
        "regions": ["美国","德国","巴西","印度","澳大利亚","荷兰"],
        "buyer_types": ["光伏安装商","EPC工程商","电力公司","光伏分销商","工商业采购商"],
        "folder": "光伏行业",
    },
    "medical_device": {
        "name_cn": "医疗器械&设备", "name_en": "Medical Devices",
        "emoji": "🏥",
        "regions": ["美国","德国","巴西","印尼","尼日利亚","墨西哥"],
        "buyer_types": ["医院集团","医疗器械分销商","GPO采购组织","诊所连锁","政府卫生采购"],
        "folder": "医疗器械行业",
    },
    "laser_equipment": {
        "name_cn": "激光设备&精密加工", "name_en": "Laser Equipment",
        "emoji": "🔩",
        "regions": ["美国","德国","日本","意大利","墨西哥","土耳其"],
        "buyer_types": ["钣金加工厂","汽车零部件厂","电子代工厂","标牌制作厂","设备分销商"],
        "folder": "激光设备行业",
    },
    "ev_charger": {
        "name_cn": "新能源汽车&充电桩", "name_en": "EV & Charging Infrastructure",
        "emoji": "🚗",
        "regions": ["德国","挪威","荷兰","泰国","巴西","阿联酋"],
        "buyer_types": ["汽车经销商","车队运营商","租车公司","充电网络运营商","加油站连锁"],
        "folder": "新能源汽车行业",
    },
    "construction_machinery": {
        "name_cn": "工程机械&农业机械", "name_en": "Construction & Ag Machinery",
        "emoji": "🏗️",
        "regions": ["印尼","印度","巴西","沙特","阿联酋","尼日利亚"],
        "buyer_types": ["建筑公司","设备租赁公司","矿业公司","大型农场","政府招标","经销商"],
        "folder": "工程机械行业",
    },
    "telecom_iot": {
        "name_cn": "通信设备&物联网", "name_en": "Telecom Equipment & IoT",
        "emoji": "📡",
        "regions": ["巴西","印尼","尼日利亚","沙特","马来西亚","泰国"],
        "buyer_types": ["电信运营商","ISP","系统集成商","智慧城市承包商","安防公司","银行/教育IT"],
        "folder": "通信设备行业",
    },
}

# ══════════════════════════════════════════════════════════
# LLM
# ══════════════════════════════════════════════════════════
def call_llm(prompt, max_tokens=4000):
    body = json.dumps({
        "model": "deepseek-chat", "messages": [{"role":"user","content":prompt}],
        "max_tokens": max_tokens, "temperature": 0.5,
    }).encode()
    req = Request("https://api.deepseek.com/chat/completions", data=body,
                  headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"})
    resp = json.loads(urlopen(req, timeout=120).read())
    return resp["choices"][0]["message"]["content"].strip()

# ══════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"rotation": list(INDUSTRIES.keys()), "idx": 0, "total": 0, "seen": [], "by_industry": {}}

def save_state(s):
    s["last_run"] = datetime.now(timezone.utc).isoformat()
    STATE_FILE.write_text(json.dumps(s, indent=2, ensure_ascii=False))

# ══════════════════════════════════════════════════════════
# GENERATION
# ══════════════════════════════════════════════════════════
def generate_buyers(industry_key, cfg, count=10):
    """Generate enhanced buyer profiles. Splits into 2 batches of 5 for JSON reliability."""
    all_buyers = []
    for batch in range(2):
        n = 5
        prompt = f"""列出{n}家真实的海外买家，可能采购中国{cfg['name_cn']}。2026年必须存续经营。

国家: {', '.join(cfg['regions'][:5])}
类型: {'; '.join(cfg['buyer_types'][:4])}

返回JSON数组，每家有这些字段:
company_name, website, founded, headquarters, employees, annual_revenue,
business_description(对外销售什么), procurement_needs(从中国采购什么),
partner_profile(找什么供应商), key_markets(主要市场),
china_presence(在华分支，没有写暂无), competitors, market_share,
recent_news(2025-2026动态), buyer_type, country, size(Enterprise/Medium/Small),
relevance(45-95), why(一句话)

只返回JSON数组，不返回markdown。字段值尽量精简。"""
        try:
            result = call_llm(prompt, max_tokens=4000)
            result = result.replace("```json","").replace("```","").strip()
            batch_buyers = json.loads(result)
            if not isinstance(batch_buyers, list):
                batch_buyers = []
            all_buyers.extend(batch_buyers)
            print(f"  Batch {batch+1}: {len(batch_buyers)} buyers")
        except Exception as e:
            print(f"  Batch {batch+1} failed: {e}")
        time.sleep(1)
    return all_buyers

def dedup_and_store(buyers, industry_key, state):
    """Dedup by company name hash, append to JSON, update state."""
    seen = set(state.get("seen", []))
    new = []
    for b in buyers:
        h = hashlib.md5(b["company_name"].lower().strip().encode()).hexdigest()
        if h not in seen:
            seen.add(h)
            b["id"] = h[:12]
            b["industry"] = industry_key
            b["generated_at"] = datetime.now(timezone.utc).isoformat()
            new.append(b)

    if not new:
        return []

    state["seen"] = list(seen)

    # Load existing
    fpath = OUTPUT_DIR / f"{industry_key}_leads.json"
    existing = []
    if fpath.exists():
        try:
            existing = json.loads(fpath.read_text()).get("leads", [])
        except: pass

    existing.extend(new)

    with open(fpath, "w") as f:
        json.dump({
            "industry": cfg["name_cn"],
            "total": len(existing),
            "updated": datetime.now(timezone.utc).isoformat(),
            "leads": existing,
        }, f, indent=2, ensure_ascii=False)

    return new


# ══════════════════════════════════════════════════════════
# BATCH UPGRADE: enrich existing 160 buyers with full profiles
# ══════════════════════════════════════════════════════════
def upgrade_existing_buyers():
    """Take existing leads (name+website only) and call DeepSeek to get full profile."""
    for ik, cfg in INDUSTRIES.items():
        fpath = OUTPUT_DIR / f"{ik}_leads.json"
        if not fpath.exists():
            continue
        data = json.loads(fpath.read_text())
        leads = data.get("leads", [])

        # Check if already upgraded (has "founded" field)
        upgraded = sum(1 for l in leads if "founded" in l and l.get("founded"))
        total = len(leads)
        if upgraded == total:
            print(f"  {cfg['emoji']} {cfg['name_cn']}: all {total} already upgraded, skip")
            continue

        print(f"  {cfg['emoji']} {cfg['name_cn']}: {upgraded}/{total} upgraded, enriching...")

        to_upgrade = [l for l in leads if "founded" not in l or not l.get("founded")]
        # Process in batches of 5
        for i in range(0, len(to_upgrade), 5):
            batch = to_upgrade[i:i+5]
            names = [b["company_name"] for b in batch]
            prompt = f"""为公司提供详细档案（每家4-5行）:

{chr(10).join([f'{j+1}. {n}' for j,n in enumerate(names)])}

返回JSON数组，每家有:
company_name(沿用原名), website, founded, headquarters, employees, annual_revenue,
business_description, procurement_needs(从中国采购{cfg['name_cn']}相关产品),
partner_profile, key_markets, china_presence, competitors, market_share,
recent_news(2025-2026), buyer_type(沿用或补充), country, size, relevance, why

只返回JSON数组。"""
            try:
                result = call_llm(prompt, max_tokens=4000)
                result = result.replace("```json","").replace("```","").strip()
                enriched = json.loads(result)
                for e in enriched:
                    name = e.get("company_name","")
                    for l in leads:
                        if l.get("company_name","").lower() == name.lower():
                            for k, v in e.items():
                                if k != "company_name" or not l.get(k):
                                    l[k] = v
                            break
                print(f"    Enriched {len(enriched)}/{len(batch)}")
            except Exception as e:
                print(f"    Batch failed: {e}")
            time.sleep(1)

        # Save back
        with open(fpath, "w") as f:
            json.dump({"industry": cfg["name_cn"], "total": len(leads), "updated": datetime.now(timezone.utc).isoformat(), "leads": leads}, f, indent=2, ensure_ascii=False)
    print("  ✅ Upgrade complete")


# ══════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  🌍 B2B Lead Miner v2")
    print(f"  {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
    print("=" * 60)

    state = load_state()

    # Step 1: Upgrade existing buyers if needed (first run only)
    if not state.get("upgraded"):
        print("\n📋 Upgrading existing buyers to full profiles...")
        upgrade_existing_buyers()
        state["upgraded"] = True
        save_state(state)

    # Step 2: Generate new buyers from next industry
    ik = state["rotation"][state["idx"]]
    cfg = INDUSTRIES[ik]
    state["idx"] = (state["idx"] + 1) % len(INDUSTRIES)

    print(f"\n📌 Industry: {cfg['emoji']} {cfg['name_cn']}")
    print(f"🧠 Generating 10 enhanced buyer profiles...")
    buyers = generate_buyers(ik, cfg, count=10)

    new = dedup_and_store(buyers, ik, state)
    state["total"] += len(new)
    state["by_industry"][ik] = state["by_industry"].get(ik, 0) + len(new)
    save_state(state)

    print(f"\n✅ Added {len(new)} new buyers to {cfg['name_cn']}")
    print(f"   Industry total: ~{state['by_industry'].get(ik, 0)}")
    print(f"   Global total: {state['total']}")

if __name__ == "__main__":
    main()
