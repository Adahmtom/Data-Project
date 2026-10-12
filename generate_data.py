"""
Builds the simulated FormWise dataset used in this project.

FormWise is a made-up online legal-document subscription service. Nothing here
is real customer data. The script is seeded, so every run produces the same
database at data/formwise.db.

Tables:
  visitors                one row per visitor (first visit, device, channel, country)
  events                  front-end analytics events, landing_view to purchase
  experiment_assignments  A/B test assignment for the paywall_v2 test
  subscriptions           one row per paid subscription
  payments                payments from the billing system
  marketing_touches       marketing touchpoints before each visitor's first visit
"""
import os, sqlite3
import numpy as np
import pandas as pd

rng = np.random.default_rng(20261009)
HERE = os.path.dirname(os.path.abspath(__file__))
START, END = pd.Timestamp("2025-01-01"), pd.Timestamp("2026-06-30")
AS_OF = pd.Timestamp("2026-09-30")
EXP_START, EXP_END = pd.Timestamp("2026-04-06"), pd.Timestamp("2026-06-14 23:59:59")

N = 180_000
days = (END - START).days + 1
# traffic grows over time and dips on weekends
day_idx = np.arange(days)
weights = (1 + day_idx / days * 0.6) * np.where(pd.date_range(START, END).dayofweek < 5, 1.15, 0.75)
visit_day = rng.choice(days, size=N, p=weights / weights.sum())
visit_ts = START + pd.to_timedelta(visit_day, "D") + pd.to_timedelta(rng.integers(0, 86400, N), "s")

device = rng.choice(["desktop", "mobile", "tablet"], N, p=[0.52, 0.41, 0.07])
channel = rng.choice(["organic_search", "paid_search", "direct", "social", "email", "referral"], N,
                     p=[0.38, 0.22, 0.15, 0.12, 0.08, 0.05])
country = rng.choice(["CA", "US", "UK", "AU"], N, p=[0.30, 0.45, 0.15, 0.10])
doc_type = rng.choice(["lease_agreement", "will", "power_of_attorney", "bill_of_sale", "nda", "promissory_note"], N,
                      p=[0.24, 0.20, 0.14, 0.16, 0.14, 0.12])

visitors = pd.DataFrame({
    "visitor_id": [f"v{i:06d}" for i in range(N)],
    "first_visit_ts": visit_ts, "device": device, "channel": channel,
    "country": country, "landing_doc_type": doc_type,
}).sort_values("first_visit_ts").reset_index(drop=True)
v = visitors

# experiment assignment
in_exp = (v.first_visit_ts >= EXP_START) & (v.first_visit_ts <= EXP_END)
variant = np.where(rng.random(N) < 0.5, "control", "treatment")
v["variant"] = np.where(in_exp, variant, None)

# funnel
p_start = np.select([v.device == "desktop", v.device == "mobile"], [0.50, 0.41], 0.46)
p_start = p_start + np.select([v.channel == "paid_search", v.channel == "social"], [0.04, -0.07], 0)
p_complete = np.select([v.device == "desktop", v.device == "mobile"], [0.66, 0.49], 0.60)
p_complete = p_complete + np.where(v.landing_doc_type == "will", -0.08, 0)
p_paywall = np.full(N, 0.95)
p_buy = np.select([v.device == "desktop", v.device == "mobile"], [0.205, 0.135], 0.17)
p_buy = p_buy + np.select([v.channel == "paid_search", v.channel == "email", v.channel == "social"],
                          [0.03, 0.04, -0.04], 0)
# treatment arm: paywall with a document preview and a plan comparison
lift = np.select([v.device == "mobile", v.device == "desktop"], [1.30, 1.10], 1.12)
p_buy = np.where(v.variant == "treatment", p_buy * lift, p_buy)

s1 = rng.random(N) < p_start
s2 = s1 & (rng.random(N) < p_complete)
s3 = s2 & (rng.random(N) < p_paywall)
s4 = s3 & (rng.random(N) < p_buy)
v["reached_template_start"], v["reached_form_complete"] = s1, s2
v["reached_paywall"], v["purchased"] = s3, s4

# events
ev = []
def add(mask, name, base_ts, lo, hi):
    idx = np.where(mask)[0]
    ts = base_ts[idx] + pd.to_timedelta(rng.integers(lo, hi, len(idx)), "s")
    ev.append(pd.DataFrame({"visitor_id": v.visitor_id.values[idx], "event_name": name, "event_ts": ts,
                            "device": v.device.values[idx]}))
    out = base_ts.copy(); out[idx] = ts
    return out
t0 = pd.Series(v.first_visit_ts.values)
ev.append(pd.DataFrame({"visitor_id": v.visitor_id, "event_name": "landing_view", "event_ts": t0, "device": v.device}))
t1 = add(s1, "template_start", t0, 20, 600)
t2 = add(s2, "form_complete", t1, 240, 2400)
t3 = add(s3, "paywall_view", t2, 2, 30)
t4 = add(s4, "checkout_start", t3, 10, 300)
t5 = add(s4, "purchase", t4, 30, 240)
events = pd.concat(ev, ignore_index=True)

# user_id is created at checkout
buyers = v[s4].copy()
buyers["user_id"] = [f"u{i:05d}" for i in range(len(buyers))]
buyers["purchase_ts"] = t5[s4].values
vid2uid = dict(zip(buyers.visitor_id, buyers.user_id))
events["user_id"] = np.where(events.event_name.isin(["checkout_start", "purchase"]),
                             events.visitor_id.map(vid2uid), None)
events["order_id"] = np.where(events.event_name == "purchase",
                              "o" + events.visitor_id.map(vid2uid).fillna("").str[1:], None)
# app version by release date
events["app_version"] = np.where(events.device == "mobile",
                                 np.where(events.event_ts < pd.Timestamp("2026-03-09"), "3.1.4",
                                          np.where(events.event_ts < pd.Timestamp("2026-05-25"), "3.2.0", "3.2.1")),
                                 "web")

# subscriptions & payments
plan = np.where(rng.random(len(buyers)) < np.where(buyers.device == "desktop", 0.42, 0.33), "annual", "monthly")
price = {"monthly": 39.99, "annual": 95.88}
docs = rng.poisson(np.where(plan == "annual", 2.6, 1.6))  # docs created in the first 30 days
buyers["plan"], buyers["docs_created_30d"] = plan, docs

subs, pays = [], []
for i, r in enumerate(buyers.itertuples(index=False)):
    start = r.purchase_ts
    engaged = r.docs_created_30d >= 2
    if r.plan == "monthly":
        n_paid, t = 1, start
        while True:
            nxt = t + pd.DateOffset(months=1)
            if nxt > AS_OF: break
            m = n_paid
            hazard = 0.42 if m == 1 else max(0.07, 0.20 * 0.85 ** (m - 2))
            hazard *= 0.62 if engaged else 1.0
            hazard *= 1.15 if r.device == "mobile" else 1.0
            if rng.random() < hazard:
                break
            t, n_paid = nxt, n_paid + 1
        cancelled = (t + pd.DateOffset(months=1)) <= AS_OF
        end = t + pd.DateOffset(months=1) if cancelled else pd.NaT
        pay_dates = [start + pd.DateOffset(months=k) for k in range(n_paid)]
    else:
        n_paid = 1
        renew_at = start + pd.DateOffset(years=1)
        renewed = renew_at <= AS_OF and rng.random() < (0.62 if engaged else 0.41)
        if renewed: n_paid = 2
        cancelled = (renew_at <= AS_OF) and not renewed
        end = renew_at if cancelled else pd.NaT
        pay_dates = [start] + ([renew_at] if renewed else [])
    subs.append((r.user_id, r.visitor_id, r.plan, start, end, n_paid, r.docs_created_30d))
    for k, d in enumerate(pay_dates):
        pays.append((f"p{len(pays):06d}", r.user_id, d, price[r.plan], "new" if k == 0 else "renewal"))

subscriptions = pd.DataFrame(subs, columns=["user_id", "visitor_id", "plan", "start_ts", "end_ts",
                                            "payments_made", "docs_created_30d"])
payments = pd.DataFrame(pays, columns=["payment_id", "user_id", "payment_ts", "amount", "payment_type"])

# marketing touches (for attribution)
touches = []
n_touch = rng.choice([1, 2, 3, 4], N, p=[0.45, 0.30, 0.17, 0.08])
chans = ["organic_search", "paid_search", "direct", "social", "email", "referral"]
for vid, ts, ch, k in zip(v.visitor_id, v.first_visit_ts, v.channel, n_touch):
    prior = rng.choice(chans, k - 1, p=[0.25, 0.20, 0.05, 0.32, 0.10, 0.08]) if k > 1 else []
    offs = sorted(rng.integers(3600, 14 * 86400, k - 1), reverse=True)
    for c, o in zip(prior, offs):
        touches.append((vid, c, ts - pd.Timedelta(seconds=int(o))))
    touches.append((vid, ch, ts))  # last touch is the visit's own channel
marketing_touches = pd.DataFrame(touches, columns=["visitor_id", "channel", "touch_ts"])

# tracking noise found in real event pipelines
purch = events[events.event_name == "purchase"]
# client retries send some purchase events twice
dups = purch.sample(frac=0.025, random_state=1).copy()
dups["event_ts"] = dups.event_ts + pd.to_timedelta(rng.integers(1, 5, len(dups)), "s")
# app 3.2.0 sometimes sends purchase events without user_id / order_id
bug = events.index[(events.event_name == "purchase") & (events.app_version == "3.2.0")]
bug = rng.choice(bug, int(len(bug) * 0.45), replace=False)
events.loc[bug, ["user_id", "order_id"]] = None
# some web purchase events never arrive (ad blockers, closed tabs)
lost = events.index[(events.event_name == "purchase") & (events.app_version == "web")]
lost = rng.choice(lost, int(len(lost) * 0.03), replace=False)
events = events.drop(index=lost)
events = pd.concat([events, dups], ignore_index=True).sort_values("event_ts").reset_index(drop=True)
events.insert(0, "event_id", [f"e{i:07d}" for i in range(len(events))])

assign = v.loc[in_exp, ["visitor_id", "variant", "first_visit_ts"]].rename(
    columns={"first_visit_ts": "assigned_ts"})
assign.insert(1, "experiment_id", "paywall_v2")

visitors_out = v[["visitor_id", "first_visit_ts", "device", "channel", "country", "landing_doc_type"]]

# write
db = os.path.join(HERE, "formwise.db")
if os.path.exists(db): os.remove(db)
con = sqlite3.connect(db)
tables = {"visitors": visitors_out, "events": events, "experiment_assignments": assign,
          "subscriptions": subscriptions, "payments": payments, "marketing_touches": marketing_touches}
for name, df in tables.items():
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[c]):
            out[c] = out[c].dt.strftime("%Y-%m-%d %H:%M:%S")
    out.to_sql(name, con, index=False)
    print(f"{name:24s} {len(out):>8,d} rows")
con.close()
