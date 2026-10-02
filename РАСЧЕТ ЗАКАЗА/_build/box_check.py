# -*- coding: utf-8 -*-
"""Вес ящика весового товара по приходам на склад РЦ.
Поставщик привозит на РЦ целыми ящиками. Ищем не НОД (его валят единичные
довесы и корректировки), а самое частое количество прихода и долю приходов,
кратных ему. Высокая доля = это и есть ящик."""
import os, csv
from google.cloud import bigquery

HERE = os.path.dirname(os.path.abspath(__file__))
os.environ.setdefault("GOOGLE_APPLICATION_CREDENTIALS",
    os.path.join(HERE, "credentials", "family-market-analytics-23fbcbcee571c.json"))
c = bigquery.Client(project="family-market-analytics")

Q = """
WITH ves AS (
  SELECT barcode, pname, unit_type, order_step
  FROM `family-market-analytics.family_market.transfer_pack_v5`
  WHERE unit_type LIKE 'ves%'
),
a AS (
  SELECT i.barcode, ROUND(i.quantity, 3) AS q
  FROM `family-market-analytics.family_market.incoming_transactions` i
  JOIN ves v USING (barcode)
  WHERE i.delivery_type = 'РЦ' AND i.quantity > 0
),
f AS (SELECT barcode, q, COUNT(*) AS c FROM a GROUP BY 1, 2),
md AS (
  SELECT barcode, q AS moda, c AS moda_n FROM (
    SELECT barcode, q, c,
           ROW_NUMBER() OVER (PARTITION BY barcode ORDER BY c DESC, q DESC) rn
    FROM f)
  WHERE rn = 1
)
SELECT a.barcode, ANY_VALUE(v.pname) AS pname, ANY_VALUE(v.unit_type) AS ut,
       ANY_VALUE(v.order_step) AS step_seychas,
       md.moda AS moda, COUNT(*) AS prihodov,
       COUNTIF(ABS(a.q / md.moda - ROUND(a.q / md.moda)) < 0.02) AS kratnyh,
       ROUND(MIN(a.q), 3) AS q_min, ROUND(MAX(a.q), 3) AS q_max,
       ROUND(AVG(a.q), 3) AS q_avg
FROM a
JOIN md USING (barcode)
JOIN ves v ON v.barcode = a.barcode
GROUP BY a.barcode, md.moda
HAVING prihodov >= 5
"""
rows = list(c.query(Q).result())
out = []
for r in rows:
    share = (r.kratnyh / r.prihodov) if r.prihodov else 0.0
    out.append(dict(barcode=r.barcode, pname=r.pname, unit_type=r.ut,
                    step_seychas=r.step_seychas, moda=r.moda,
                    prihodov=r.prihodov, kratnyh=r.kratnyh,
                    dolya_kratnyh=round(share, 3),
                    q_min=r.q_min, q_max=r.q_max, q_avg=r.q_avg))
out.sort(key=lambda d: (-d["dolya_kratnyh"], -d["prihodov"]))
p = os.path.join(HERE, "box_check.csv")
with open(p, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(out[0].keys()), delimiter=";")
    w.writeheader(); w.writerows(out)
print("gotovo:", len(out), "strok")
