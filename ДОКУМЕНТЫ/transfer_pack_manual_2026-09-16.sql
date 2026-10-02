-- 0. Схема: убедиться, что нет обязательных столбцов кроме этих трёх
SELECT column_name, data_type, is_nullable
FROM `family-market-analytics.family_market.INFORMATION_SCHEMA.COLUMNS`
WHERE table_name = 'transfer_pack_manual'
ORDER BY ordinal_position;

-- 1. Кофейные шоубоксы: шаг подтверждён числом в скобках названия.
--    Явный order_step снимает WARN "Шаг взят из названия" (5 артикулов).
MERGE `family-market-analytics.family_market.transfer_pack_manual` T
USING (
  SELECT * FROM UNNEST([
    STRUCT('2938080084499' AS barcode, 'sht_nedelimyy' AS unit_type, 20.0 AS order_step),
    ('2938080084505', 'sht_nedelimyy', 20.0),   -- МакКофе Голд 3в2 16г (20)
    ('8887290109932', 'sht_nedelimyy', 20.0),   -- МакКофе Арабіка 3в2 16г (20)
    ('8887290109956', 'sht_nedelimyy', 20.0),   -- МакКофе Голд 3в2 16г (20)
    ('8711000609415', 'sht_nedelimyy', 30.0)    -- Maxwell House Origin 18г, шоубокс из артикула
  ])
) S
ON T.barcode = S.barcode
WHEN MATCHED THEN UPDATE SET unit_type = S.unit_type, order_step = S.order_step
WHEN NOT MATCHED THEN INSERT (barcode, unit_type, order_step)
                      VALUES (S.barcode, S.unit_type, S.order_step);

-- 2. Якобз Латте 13гр (min_q=23, n=170) и Нескафе Айріш 14гр (min_q=19, n=22).
--    Штрих-кодов в отчёте нет - найти:
SELECT barcode, pname, unit_type, order_step, min_q, n
FROM `family-market-analytics.family_market.transfer_pack_v5`
WHERE LOWER(pname) LIKE '%латте 13%' OR LOWER(pname) LIKE '%айріш 14%';

-- 3. Проверка шоубокса по приходам на РЦ, прежде чем вносить 24 и 20.
--    Логика как в box_check: не НОД, а самое частое количество прихода
--    и доля приходов, кратных ему. Доля выше ~0.8 = это и есть шоубокс.
WITH a AS (
  SELECT barcode, ROUND(quantity) AS q
  FROM `family-market-analytics.family_market.incoming_transactions`
  WHERE delivery_type = 'РЦ' AND quantity > 0
    AND barcode IN ('ПОДСТАВИТЬ_ИЗ_ЗАПРОСА_2')
),
f AS (SELECT barcode, q, COUNT(*) c FROM a GROUP BY 1,2),
md AS (
  SELECT barcode, q AS moda FROM (
    SELECT barcode, q, ROW_NUMBER() OVER (PARTITION BY barcode ORDER BY c DESC, q DESC) rn
    FROM f) WHERE rn = 1
)
SELECT a.barcode, md.moda, COUNT(*) AS prihodov,
       ROUND(COUNTIF(MOD(a.q, md.moda) = 0) / COUNT(*), 3) AS dolya_kratnyh
FROM a JOIN md USING (barcode)
GROUP BY 1, 2;

-- 4. Пепсі-Кола 0,33 і 1л (min_q=12): похоже на термоусадку 12,
--    автоматом не вносить, подтвердить на РЦ.
-- 5. Пакет Фемелі і Запальничка Кліппер: не трогать - у них min_q спорит
--    с жёсткими правилами BAG_BLOCK=50 и блоком 5, правила главнее.
