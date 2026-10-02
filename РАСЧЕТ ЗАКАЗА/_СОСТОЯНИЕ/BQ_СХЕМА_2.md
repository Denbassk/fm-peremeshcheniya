# BigQuery разведка 2: family-market-analytics.family_market
снято 2026-09-22 17:09

## _dash_dead
  колонки: barcode:STRING; p:STRING; c:STRING; s:STRING; store:STRING; qty:FLOAT64; value:FLOAT64
  строк: 8924
  точек по store: 38
  топ: Ювілейний 67(338); Танкопія 16(337); Полевая магазин(337); Героїв Харкова 160(314); Переяславська 23(302); Зубенка 31В5(292); Бучми 32Б1(288); Олімпійська 9А(285); Роганська 148(272); Валентинівська 24Б(269); Бучми 32(267); Богдана Хмельницького 8(263)
  пример 1: barcode=7638900227550; p=Батарейки Енерджайзер Еврідей ААА 4 ; c=Батарейки; s=Діст Сістем; store=Гвардійців Широнінців 54; qty=3.0; value=116.64000000000001
  пример 2: barcode=7638900227550; p=Батарейки Енерджайзер Еврідей ААА 4 ; c=Батарейки; s=Діст Сістем; store=Ювілейний 67; qty=3.0; value=116.64000000000001

## _dash_s90
  колонки: barcode:STRING; store:STRING; nm:STRING; qty90:FLOAT64
  строк: 64130
  точек по store: 38
  топ: Полевая магазин(2002); Михайля Семенка 17(1926); Іскринський 19В(1909); Героїв Небесної Сотні 14/1(1899); Шевченко 341(1879); Салтівське шосе 264В(1844); Астрономічна 44Г(1844); Грозненська 38(1818); Байрона 156(1782); Валентинівська 50А(1769); Бучми 52(1764); Ньютона 102(1758)
  пример 1: barcode=132303198020; store=Грозненська 38; nm=Кузя Сервіс Пакет Суп.Люкс Д/Сміт 60; qty90=1.0
  пример 2: barcode=132303198020; store=Ньютона 102; nm=Кузя Сервіс Пакет Суп.Люкс Д/Сміт 60; qty90=4.0

## _dash_stk
  колонки: barcode:STRING; store:STRING; qty:FLOAT64; qty_pos:FLOAT64; n_pos:INT64; nm:STRING; cost:FLOAT64
  строк: 60411
  точек по store: 38
  топ: Полевая магазин(1870); Михайля Семенка 17(1743); Іскринський 19В(1728); Астрономічна 44Г(1726); Салтівське шосе 264В(1726); Петра Григоренка 37(1707); Героїв Небесної Сотні 14/1(1701); Переяславська 23(1669); Роганська 148(1662); Бучми 32Б1(1659); Шевченко 341(1656); Бучми 52(1655)
  пример 1: barcode=19230010208; store=Полевая магазин; qty=2.134; qty_pos=2.134; n_pos=1; nm=Роганський МК Сардельки Богатирські ; cost=97.14
  пример 2: barcode=19230010208; store=Байрона 138/1; qty=0.324; qty_pos=0.324; n_pos=1; nm=Роганський МК Сардельки Богатирські ; cost=97.14

## _dash_a26
  колонки: mo:INT64; store:STRING; barcode:STRING; product_name:STRING; category:STRING; supplier:STRING; in_matrix:BOOL; cul:BOOL; qty:FLOAT64; rev:FLOAT64; gp:FLOAT64
  строк: 451340
  пример 1: mo=1; store=Іскринський 19В; barcode=4820004380429; product_name=Ігристе Французький Бульвар брют біл; category=Шампанское / Игристое; supplier=БІР Алкоголь; in_matrix=True; cul=False; qty=3.0; rev=490.5; gp=78.339924
  пример 2: mo=1; store=Ювілейний 67; barcode=4820004380429; product_name=Ігристе Французький Бульвар брют біл; category=Шампанское / Игристое; supplier=БІР Алкоголь; in_matrix=True; cul=False; qty=1.0; rev=163.5; gp=1.3400039999999933

## _dash_r26
  колонки: mo:INT64; store:STRING; category:STRING; supplier:STRING; dow:INT64; hr:INT64; cul:BOOL; receipts:INT64; revenue:FLOAT64
  строк: 60335
  пример 1: mo=None; store=None; category=None; supplier=Сік Ранок; dow=None; hr=None; cul=None; receipts=6446; revenue=230548.0
  пример 2: mo=None; store=None; category=None; supplier=Хладік; dow=None; hr=None; cul=None; receipts=108656; revenue=10029711.0

## store_canon
  колонки: store_raw:STRING; store_canon:STRING; kind:STRING; opened_from:DATE; closed_after:DATE
  строк: 49
  opened_from: 2026-01-25 .. 2026-06-10
  пример 1: store_raw=Астрономічна 44Г; store_canon=Астрономічна 44Г; kind=retail; opened_from=None; closed_after=None
  пример 2: store_raw=Полевая; store_canon=Полевая-Склад; kind=sklad; opened_from=None; closed_after=None

## store_mapping
  колонки: store_original:STRING; store_normalized:STRING; is_active:BOOL
  строк: 43
  пример 1: store_original=Полевая 83; store_normalized=Полевая 83; is_active=False
  пример 2: store_original=Професорська 12; store_normalized=Професорська 12; is_active=False

## store_info
  колонки: store:STRING; opened_date:DATETIME; last_transaction:DATETIME; months_active:INT64; days_active:INT64
  строк: 37
  opened_date: 2025-01-01 08:04:59 .. 2025-11-12 08:07:07
  точек по store: 37
  топ: Полевая 83(1); Нескорених 33(1); Професорська 12(1); Іскрінський 19(1); Михайля Семенка 17(1); Шевченко 341(1); Небесної Сотні 14/1(1); Переяславська 23(1); Ньютона 111(1); Амосова 5А(1); Героїв Праці 33(1); Пр-т Тракторобудiвникiв 95(1)
  пример 1: store=Михайля Семенка 17; opened_date=2025-11-12 08:07:07; last_transaction=2025-11-30 21:59:34; months_active=1; days_active=19
  пример 2: store=Нескорених 33; opened_date=2025-11-01 07:37:16; last_transaction=2025-11-30 21:44:32; months_active=1; days_active=30

## assortment_matrix_full
  колонки: barcode:STRING; supplier:STRING; product_name:STRING; category:STRING; decision:STRING; cost:FLOAT64; profit:FLOAT64; coverage:FLOAT64; status:STRING
  строк: 2046
  пример 1: barcode=4823127313923; supplier=Авангард; product_name=Авангард Сан Санич Насіння Преміум С; category=Семечки; decision=ОСТАВИТЬ; cost=34.86; profit=74317.14; coverage=38.0; status=🟢 Прибыльный
  пример 2: barcode=4823127314494; supplier=Авангард; product_name=Авангард Сан Санич Насіння Преміум С; category=Семечки; decision=ОСТАВИТЬ; cost=34.86; profit=50082.28; coverage=38.0; status=🟢 Прибыльный

## stock_matrix
  колонки: record_id:STRING; barcode:STRING; product_name:STRING; store:STRING; quantity:FLOAT64; price_retail:FLOAT64; snapshot_date:DATE; source_file:STRING; loaded_at:DATETIME; supplier:STRING; cost_price:FLOAT64; article:STRING
  строк: 244507
  snapshot_date: 2026-01-08 .. 2026-09-18
  пример 1: record_id=123456777_Полевая-Склад_20260918; barcode=123456777; product_name=Пергамент Маестро Без Втулки 6м; store=Полевая-Склад; quantity=1.0; price_retail=10.5; snapshot_date=2026-09-18; source_file=Остатки на складах на 18.09.2026.xls; loaded_at=2026-09-18 10:40:23.154976; supplier=None; cost_price=9.6; article=None
  пример 2: record_id=132303198020_Іскрінський_19_20260918; barcode=132303198020; product_name=Кузя Сервіс Пакет Суп.Люкс Д/Сміт 60; store=Іскрінський 19; quantity=3.0; price_retail=16.2; snapshot_date=2026-09-18; source_file=Остатки на складах на 18.09.2026.xls; loaded_at=2026-09-18 10:40:23.154976; supplier=None; cost_price=10.77; article=None

## rotation_candidates
  колонки: supplier:STRING; barcode:STRING; product_name:STRING; product_prefix:STRING; in_active_matrix:BOOL; first_arrival_date:DATE; last_arrival_date:DATE; network_first_date:DATE; snapshot_date:DATE; age_days:INT64; age_months:FLOAT64; age_days_network:INT64; real_coverage:INT64; stale_in_stores:INT64; stale_stores_with_stock:INT64; stale_stores_with_any_sales:INT64; stores_with_stock:INT64; stores_with_sales:INT64; total_stores:INT64; stale_ratio_pct:FLOAT64; stores_stock_qty:FLOAT64; central_stock_qty:FLOAT64; expired_stock_qty:FLOAT64; total_cost_frozen:FLOAT64; total_retail_frozen:FLOAT64; sales_period_torgsoft:FLOAT64; sales_total_ytd:FLOAT64; sales_revenue_ytd:FLOAT64; profit_2026:FLOAT64; last_sale_date:DATE; days_since_last_sale:INT64; period_start:DATE; profit_in_period:FLOAT64; sales_qty_in_period:FLOAT64; sales_revenue_in_period:FLOAT64; is_weighted:BOOL; sales_jan:FLOAT64; rev_jan:FLOAT64; sales_feb:FLOAT64; rev_feb:FLOAT64; sales_mar:FLOAT64; rev_mar:FLOAT64; sales_apr:FLOAT64; rev_apr:FLOAT64; sales_may:FLOAT64; rev_may:FLOAT64; calculated_at:DATE; period_days:INT64; profit_annual_forecast:FLOAT64; sales_units_annual_forecast:FLOAT64; sales_revenue_annual_forecast:FLOAT64; profit_gap:FLOAT64; rotation_status:STRING
  строк: 3536
  first_arrival_date: 2025-01-02 .. 2026-06-05
  пример 1: supplier=ЄВРОМІКС; barcode=4820108004962; product_name=Презервативи Reflex Латексні з Сил.з; product_prefix=Презервативи; in_active_matrix=False; first_arrival_date=2025-06-08; last_arrival_date=2025-06-08; network_first_date=2025-06-08; snapshot_date=2026-06-09; age_days=366; age_months=12.2; age_days_network=366; real_coverage=1; stale_in_stores=1; stale_stores_with_stock=1; stale_stores_with_any_sales=1
  пример 2: supplier=ЄВРОМІКС; barcode=5052197053401; product_name=Презервативи Durex Dual Extase 3шт; product_prefix=Презервативи; in_active_matrix=False; first_arrival_date=2025-06-08; last_arrival_date=2025-06-08; network_first_date=2025-06-08; snapshot_date=2026-06-09; age_days=366; age_months=12.2; age_days_network=366; real_coverage=0; stale_in_stores=1; stale_stores_with_stock=1; stale_stores_with_any_sales=0

## stale_goods_raw
  колонки: product_name:STRING; barcode:STRING; sales_qty:FLOAT64; stock_qty:FLOAT64; store:STRING; cost_per_unit:FLOAT64; cost_sum:FLOAT64; retail_price:FLOAT64; discount_price:FLOAT64; retail_sum:FLOAT64; snapshot_date:DATE; loaded_at:TIMESTAMP
  строк: 14026
  snapshot_date: 2026-06-09 .. 2026-06-09
  точек по store: 42
  топ: Зернова 6/5(1084); Полевая-Магазин(910); Полевая-Просрок(678); Роганська 148(454); Полевая-Склад(452); Нескорених 4Д(438); Пр-т Героїв Харкова 160(390); Танкопія 16(382); Зубенко 31В/5(373); Петра Григоренка 37(366); Пр-т Ювілейний 67(365); Краснодарська 171з(358)
  пример 1: product_name=Абрикос Сушений (Курага) (н.1) 500г; barcode=4820232571125; sales_qty=0.0; stock_qty=6.0; store=Полевая-Просрок; cost_per_unit=118.92; cost_sum=713.52; retail_price=333.0; discount_price=333.0; retail_sum=1998.0; snapshot_date=2026-06-09; loaded_at=2026-06-09 18:53:44.209524+00:00
  пример 2: product_name=Авангард Багет Флінт Вер. Соус Зелен; barcode=4820182746727; sales_qty=0.0; stock_qty=5.0; store=Зернова 6/5; cost_per_unit=23.28; cost_sum=116.4; retail_price=29.6; discount_price=29.6; retail_sum=148.0; snapshot_date=2026-06-09; loaded_at=2026-06-09 18:53:44.209524+00:00

## item_lifecycle_decisions
  колонки: barcode:STRING; status:STRING; decided_at:DATE; decided_by:STRING; note:STRING; revenue_at_decision:FLOAT64
  строк: 47
  decided_at: 2026-09-08 .. 2026-09-19
  пример 1: barcode=4820050234660; status=Додано; decided_at=2026-09-10; decided_by=denba; note=додано через matrix_ui; revenue_at_decision=0.0
  пример 2: barcode=4820029432240; status=Додано; decided_at=2026-09-10; decided_by=denba; note=додано через matrix_ui; revenue_at_decision=0.0

## barcode_recode_map
  колонки: old_barcode:STRING; new_barcode:STRING; product_name:STRING; note:STRING
  строк: 180
  пример 1: old_barcode=4820139280274; new_barcode=4820139280779; product_name=; note=
  пример 2: old_barcode=4820139280267; new_barcode=4820139280786; product_name=; note=

## order_reference
  колонки: barcode:STRING; pname:STRING; main_supplier:STRING; n_lines:INT64; min_q:FLOAT64; unit_type:STRING; order_step:INT64; pack_rc_zakupka:INT64; pack_source:STRING
  строк: 2373
  пример 1: barcode=4820193034523; pname=Оболонь Напій С/а Джин Грейпфрут 0,3; main_supplier=Союз (Оболонь); n_lines=1037; min_q=6.0; unit_type=sht; order_step=1; pack_rc_zakupka=6; pack_source=A_podtverzhdeno
  пример 2: barcode=4820097892069; pname=Шейки Вода Природне Джерело 0,5л Нег; main_supplier=Арсенал ПК (Шейк); n_lines=542; min_q=6.0; unit_type=sht; order_step=1; pack_rc_zakupka=6; pack_source=pravilo_postavshika

## pack_reference
  колонки: barcode:STRING; pname:STRING; art_pack:INT64; k_gcd:INT64; share_gcd:FLOAT64; nonround_ok:INT64; n:INT64; min_q:INT64; mode_q:INT64; uq:INT64; tier:STRING; pack_supplier:INT64
  строк: 2171
  пример 1: barcode=2984670043759; pname=None; art_pack=None; k_gcd=2; share_gcd=1.0; nonround_ok=264; n=351; min_q=8; mode_q=16; uq=15; tier=A_podtverzhdeno; pack_supplier=2
  пример 2: barcode=4820250942754; pname=БІР Пиво Старопрамен Світле 0,48л з/; art_pack=24; k_gcd=2; share_gcd=0.981; nonround_ok=1563; n=1608; min_q=8; mode_q=12; uq=31; tier=C_norma; pack_supplier=2

## weight_reference
  колонки: barcode:STRING; pname:STRING; n:INT64; min_q:FLOAT64; w_name:FLOAT64; is_range:BOOL; unit_kg:FLOAT64; rel_res:FLOAT64; unit_type:STRING; src_check:STRING
  строк: 156
  пример 1: barcode=2984670081225; pname=Овочі/Фрукти Капуста Пекінська 1кг; n=25; min_q=0.41; w_name=None; is_range=False; unit_kg=0.41; rel_res=0.2107; unit_type=ves_chistyy; src_check=только накладные
  пример 2: barcode=2984670061609; pname=Овочі/Фрукти Капуста 1кг; n=94; min_q=0.67; w_name=None; is_range=False; unit_kg=0.67; rel_res=0.2528; unit_type=ves_chistyy; src_check=только накладные

## abc_analysis
  колонки: barcode:STRING; product_name:STRING; supplier:STRING; name_norm:STRING; revenue:FLOAT64; first_arrival_date:DATE; age_days:INT64; share_pct:FLOAT64; share_cum_pct:FLOAT64; abc_category:STRING; calculated_at:DATE
  строк: 2000
  first_arrival_date: 2025-01-02 .. 2026-06-04
  пример 1: barcode=2984670064358; product_name=Риба Шпроти г/к 1кг; supplier=Риба; name_norm=риба шпроти г/к 1кг; revenue=14525.895799999997; first_arrival_date=2025-09-17; age_days=265; share_pct=0.004280453450320961; share_cum_pct=99.41899383770836; abc_category=C; calculated_at=2026-06-09
  пример 2: barcode=4820250942082; product_name=БІР Пиво Львівське 1715 0,48л з/б  Б; supplier=БІР; name_norm=бір пиво львівське 1715 0,48л з/б б/; revenue=219604.42; first_arrival_date=2025-01-04; age_days=521; share_pct=0.06471246319244103; share_cum_pct=70.03809202971219; abc_category=A; calculated_at=2026-06-09

## pozycii_poza_matryceyu
  колонки: barcode:STRING; nazva:STRING; pershyi_prodazh:DATE; ostannii_prodazh:DATE; chekiv_90d:INT64; chekiv_perede_90d:INT64; magazyniv_90d:INT64; magazyniv_perede_90d:INT64; vyruchka_90d:FLOAT64; stan:STRING; rishennia:STRING
  строк: 1136
  pershyi_prodazh: 2025-01-01 .. 2026-08-23
  пример 1: barcode=2938080084499; nazva=Кава Мак Кофе Арабіка 3в2 16г (20) 1; pershyi_prodazh=2026-08-18; ostannii_prodazh=2026-08-21; chekiv_90d=7; chekiv_perede_90d=0; magazyniv_90d=2; magazyniv_perede_90d=0; vyruchka_90d=72.0; stan=Власне виробництво; rishennia=Не додавати
  пример 2: barcode=2938080082242; nazva=Овочі/Фрукти Кріп 1кг з уцінкою-50%; pershyi_prodazh=2026-05-29; ostannii_prodazh=2026-05-29; chekiv_90d=1; chekiv_perede_90d=0; magazyniv_90d=1; magazyniv_perede_90d=0; vyruchka_90d=35.1; stan=Власне виробництво; rishennia=Не додавати

## transfer_transactions
  колонки: line_id:STRING; transfer_id:STRING; line_number:INT64; product_name:STRING; barcode:STRING; sender:STRING; store:STRING; quantity:FLOAT64; price_retail:FLOAT64; price_purchase:FLOAT64; amount_retail:FLOAT64; amount_purchase:FLOAT64; transfer_datetime:DATETIME; source_file:STRING; loaded_at:DATETIME; doc_date:DATE; doc_number:INT64; created_at:DATETIME
  строк: 517251
  transfer_datetime: 2025-01-02 11:22:02 .. 2026-08-31 00:00:00
  пример 1: line_id=2026-01-02|9625|4820001830064|6.000|; transfer_id=DOC_20260102_9625; line_number=1; product_name=Вода мін.Лужанська 1,5л; barcode=4820001830064; sender=Полевая 83; store=Олімпійська 9А; quantity=6.0; price_retail=43.5; price_purchase=34.56; amount_retail=261.0; amount_purchase=207.36; transfer_datetime=2026-01-02 00:00:00; source_file=Внутренние перемещения янв-март 2026; loaded_at=2026-09-04 18:01:17.602860; doc_date=2026-01-02
  пример 2: line_id=2026-01-02|9625|4823122200082|10.000; transfer_id=DOC_20260102_9625; line_number=2; product_name=Грин Дей Горілка Зелений День 0,1л; barcode=4823122200082; sender=Полевая 83; store=Олімпійська 9А; quantity=10.0; price_retail=34.1; price_purchase=26.4; amount_retail=341.0; amount_purchase=264.0; transfer_datetime=2026-01-02 00:00:00; source_file=Внутренние перемещения янв-март 2026; loaded_at=2026-09-04 18:01:17.602860; doc_date=2026-01-02

## Чистота store в turnover_monthly
  всего значений store: 3751; короче 40 символов: 3751
  топ-15: Іскрінський 19(40876); Салтівське шосе 264В(40695); Шевченко 341(40113); Астрономічна 44Г(39850); Петра Григоренка 37(38989); Пр-т Героїв Харкова 160(38927); Роганська 130/4(38548); Бучми 52(38522); Амосова 5А(38434); Валентинівська 50А(38424); Ньютона 102(38056); Качанівська 19(38004); Бучми Джерело(37975); Грозненська 38(37547); Краснодарська 171з(37498)
  редкие (похоже на мусор): -97.83(1); -17.1(1); -0.8(1); -34(1); -6.85(1); -24.11(1); -1.91(1); -3.44(1); -7.61(1); -29.14(1)
  последние периоды: 2026-08(73305); 2026-07(72946); 2026-06(72606); 2026-05(71362); 2026-04(82751); 2026-03(67557)
