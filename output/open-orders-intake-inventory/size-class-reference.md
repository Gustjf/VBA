# Bobcat MODEL_GROUP to Size Class: Reference Table

Since macro version 2.0 the size classes live in a separate size class workbook (Config `SizeClassFile`, the master list); this table is the source of the starter list `size-class-master.xlsx` (made by `make_size_class_master.py`), whose Notes column flags every BOUNDARY and CONFLICT model below with "CHECK: on a class boundary". It was first written as the companion to the version 1.x `SizeClassMap` sheet. Compiled 2026-10-09 by a research agent of the Stronger AI review from web search extracts of Bobcat regional pages, dealer spec sheets and trade press (the agent could not open the pages directly, so figures are from search snippets and should be confirmed against a current Bobcat spec sheet before the table is treated as authoritative). Every row the agent could not source is marked Low and should be checked by a person. A second check on 2026-10-10 (three research agents using web search; bobcat.com pages could not be opened directly, so figures again come from search results) confirmed the class of every row it re-checked (the Med, Low and BOUNDARY rows not fixed by the sample), found conflicting sources for T190, S530, S550, S570 and S86-2, and added T200, T250 and T320; see the Verification section.

## Rules

- **CTL / SSL** (compact track loaders, skid-steer loaders): bucket by **rated operating capacity (ROC) in lb**, lower bound inclusive, upper bound exclusive. CTL ROC is published at 35% of tipping load, SSL at 50%. Buckets: `0<2000`, `2000<2300`, `2300<2600`, `2600<3000`, `3000<3500`, `3500&Over`.
- **MX** (mini excavators): bucket by **nominal tonnes = model number / 10**, lower bound exclusive, upper bound inclusive. Buckets: `>00-01`, `>01-02`, `>02-03`, `>03-04`, `>04-05`, `>05-06`, `>06-10`, `Over10`. Evidence that the rule is nominal and not actual weight: the sample report puts E10 (actual 1,176 kg) in `>00-01` and E20 (actual about 1,950 kg) in `>01-02`, and E60 in `>05-06`.
- **GT** = matches the ground truth taken from the hand-built sample report. **BOUNDARY** = a published figure sits within 100 lb of a class edge, or regional figures straddle one.
- "-2" models are the R2-series version of the base model and take the same class.

## Compact track loaders

| MODEL_GROUP | ROC (lb) and notes | Size class | Confidence |
|---|---|---|---|
| T190 (legacy) | 1,900 (Boom & Bucket, listings with 5,430 lb tipping); one LECTURA "T190 ACS" entry 1,088 kg = 2,399. CONFLICT | 0<2000 | Med (verify) |
| T450 | 1,400 (635 kg); a third-party page shows 1,490. GT | 0<2000 | High |
| T550 | 1,995 (Bobcat NA); later model years show 2,027. BOUNDARY. GT | 0<2000 | High for the class per GT; spec is model-year sensitive |
| T590 | 2,100; metric 992 kg = 2,187 | 2000<2300 | High |
| T595 | 2,100 (Bobcat); trade press 2,200; one dealer 2,280. BOUNDARY. GT | 2000<2300 | High |
| T62 | 2,150; metric 1,025 kg = 2,260. BOUNDARY | 2000<2300 | Med |
| T630 (legacy) | 2,100 to 2,230. BOUNDARY | 2000<2300 | Med |
| T64 | 2,300 exactly. BOUNDARY | 2300<2600 | High |
| T64-2 | 2,300 exactly. BOUNDARY. GT | 2300<2600 | High |
| T650 | 2,570 (1,166 kg); some counterweight configs up to 2,675. BOUNDARY. GT | 2300<2600 | High |
| T66 | 2,450 (1,111 to 1,124 kg) | 2300<2600 | High |
| T66-2 | 2,450. GT | 2300<2600 | High |
| T76 | 2,900 (1,315 kg). BOUNDARY | 2600<3000 | High |
| T76-2 | 2,900. BOUNDARY. GT | 2600<3000 | High |
| T300 (legacy) | 3,000 exactly. BOUNDARY | 3000<3500 | Med |
| T740 | 3,200 (ISO); some dealers 3,310. GT | 3000<3500 | High |
| T770 | 3,475 (dealers, trade press); Bobcat international 1,618 kg = 3,567. BOUNDARY. GT | 3000<3500 | Med (class per GT) |
| T7X (electric) | 3,056 current; launch coverage 2,900 to 2,950. BOUNDARY | 3000<3500 | Med |
| T86 | 3,800; international 1,777 kg = 3,917 | 3500&Over | High |
| T86-2 | 3,800; one 2026 dealer 3,700. GT | 3500&Over | High |
| T870 | 3,650 (Bobcat NA); LA page 3,525; 2010 launch 3,325. BOUNDARY | 3500&Over | Med |
| T200 (legacy) | 2,000 (RitchieSpecs ISO 14397; 2001 listing at 35% of 5,715 lb tipping; LECTURA 907 kg). BOUNDARY | 2000<2300 | Med |
| T250 (legacy) | 2,500 (dealer auction listings at 35% of 7,143 lb tipping); LECTURA 1,168 kg = 2,575. BOUNDARY | 2300<2600 | Med |
| T320 (legacy) | 3,100 (Bobcat non-current models page, ISO); 3,200 (trade press); KHL 1,461 kg = 3,221. BOUNDARY | 3000<3500 | Med |

## Skid-steer loaders

| MODEL_GROUP | ROC (lb) and notes | Size class | Confidence |
|---|---|---|---|
| S70 | 760 (older listings 700) | 0<2000 | High |
| S185 (legacy) | 1,850 | 0<2000 | Med |
| S450 | 1,370 (older 1,300) | 0<2000 | High |
| S510 | 1,876 (851 kg); other years 1,650 to 1,790 | 0<2000 | High |
| S530 | 1,916 (869 kg ISO, Bobcat LatAm/SEA); EU Stage V 911 kg = 2,008. CONFLICT across the 2,000 edge | 0<2000 | Med (verify) |
| S550 | 2,041 (926 kg, 68.9 hp variant) vs 1,971 (894 kg, 61 hp variant) and 1,750 (dealers): the variant decides. CONFLICT | 0<2000 (provisional) | Low-Med (verify) |
| S570 | 1,950 SAE (dealers, spec databases); Bobcat regional ISO 944 kg = 2,081. CONFLICT | 0<2000 | Med (verify) |
| S205 (legacy) | 2,050. BOUNDARY | 2000<2300 | Med |
| S590 | 2,100; metric 953 to 998 kg. BOUNDARY | 2000<2300 | High |
| S595 | 2,200 (998 kg). BOUNDARY | 2000<2300 | Med-High |
| S62 | 2,100; OC page 1,006 kg = 2,218; 2,250 with counterweight. BOUNDARY | 2000<2300 | Med |
| S630 (legacy) | 2,180 | 2000<2300 | Med-High |
| S64 | 2,300 exactly; 2,450 with 200 lb counterweight. BOUNDARY | 2300<2600 | High |
| S64-2 | 2,300 (dealer R2 sheet). BOUNDARY | 2300<2600 | Med |
| S66 | 2,400; KR/OC 1,166 kg = 2,570. BOUNDARY | 2300<2600 | High |
| S66-2 | 2,400; OC page 1,140 kg = 2,513. BOUNDARY | 2300<2600 | Med-High |
| S650 | 2,690 (tip 5,380). BOUNDARY | 2600<3000 | High |
| S76 | 2,900 (US ISO); 3,050 with counterweight. BOUNDARY | 2600<3000 | Med-High |
| S300 (legacy) | 3,000 (listings; LECTURA 1,386 kg = 3,056). BOUNDARY | 3000<3500 (provisional) | Med |
| S330 (legacy) | 3,300 | 3000<3500 | Med-High |
| S740 | 3,100 (tip 6,200). BOUNDARY | 3000<3500 | High |
| S750 (legacy) | 3,200 | 3000<3500 | High |
| S770 | 3,350 (1,520 kg); APAC/LA 1,569 kg = 3,459; Bobcat IN 1,520 kg = 3,351; LA/SEA 1,569 kg = 3,459. BOUNDARY | 3000<3500 | High |
| S86 | 3,400 (105 hp); counterweight options 3,550 to 3,700; Bobcat OC 1,564 kg = 3,448. BOUNDARY | 3000<3500 | Med-High |
| S86-2 | 3,500 (Bobcat NA "Pro S86-2", 7,000 lb tipping); Bobcat metric 1,587 kg = 3,499; one stale dealer sheet 3,400. CONFLICT on the edge (class stays 3500&Over, the NA figure governs) | 3500&Over | Low-Med |
| S850 | 3,950 (launch 3,875) | 3500&Over | High |

## Mini excavators (nominal tonnes = model number / 10)

| MODEL_GROUP | Nominal t (actual weight) | Size class | Confidence |
|---|---|---|---|
| E08 | 0.8 (1,138 kg) | >00-01 | High |
| E10 | 1.0 (1,176 kg). GT | >00-01 | High |
| E10e | 1.0 (about 1,201 kg) | >00-01 | Med |
| E17 | 1.7 (1,711 kg). GT | >01-02 | High |
| E17z | 1.7 (1,712 to 1,749 kg) | >01-02 | High |
| E19 / E19e | 1.9 (1,885 kg) | >01-02 | High |
| E20 | 2.0, upper bound inclusive (about 1,953 kg). GT | >01-02 | High |
| E20z | 2.0 (1,900 to 2,019 kg) | >01-02 | High |
| E26 | 2.6 (2,950 kg). GT | >02-03 | High |
| E27z | 2.7 (2,595 to 2,705 kg) | >02-03 | High |
| E30 | 3.0, upper bound inclusive (2,950 to 3,069 kg) | >02-03 | Med |
| E32 / E32i | 3.2 (about 3,329 kg). GT | >03-04 | High |
| E34 | 3.4 (3,476 kg) | >03-04 | High |
| E35 | 3.5 (about 3,492 kg; 25 hp and 33 hp variants). GT | >03-04 | High |
| E35z | 3.5 (3,362 to 3,499 kg) | >03-04 | High |
| E37 | 3.7 (3,630 to 3,738 kg) | >03-04 | High |
| E40 | 4.0 nominal, upper bound inclusive; actual 9,180 lb = 4.16 t | >03-04 | Med (flag) |
| E42 | 4.2 (4,459 kg) | >04-05 | High |
| E45 | 4.5 (4,571 to 4,634 kg) | >04-05 | High |
| E48 | 4.8 (about 9,830 lb). GT | >04-05 | High |
| E50 | 5.0, upper bound inclusive (4,843 to 5,151 kg) | >04-05 | Med-High |
| E50z | 5.0 (4,819 kg) | >04-05 | Med-High |
| E55 | 5.5 (5,585 kg). GT | >05-06 | High |
| E55z | 5.5 (5,218 kg) | >05-06 | High |
| E55W (wheeled) | 5.5 (5,550 kg) | >05-06 | Med-High |
| E60 | 6.0, upper bound inclusive (5,586 kg). GT | >05-06 | High |
| E63 | 6.3 (6,250 kg) | >06-10 | High |
| E85 | 8.5 (8,509 to 8,608 kg) | >06-10 | High |
| E88 | 8.8 (8,919 to 9,006 kg). GT | >06-10 | High |
| E145 (large excavator) | 14.5 (15,870 kg) | Over10 | High |
| E165 (large excavator) | 16.5 (17,407 kg) | Over10 | High |

Not found: any S76-2, S740-2 or S770-2; any Bobcat mini excavator between E88 and E145.

## Verification, 2026-10-10

Each model was checked by web search, Bobcat NA pages were preferred, figures are standard ratings without counterweight, kg were converted at 2.20462 lb/kg, and bobcat.com could not be opened, so figures come from search results.

| MODEL_GROUP | Figure found (lb) | Source kind | Verdict |
|---|---|---|---|
| T62 | 2,150 (Bobcat NA) | Bobcat | SAME |
| T630 | 2,230 (35% of 6,371 tipping) | trade, spec database | SAME |
| T64 | 2,300 (Bobcat NA) | Bobcat, dealers | SAME (on the edge) |
| T76 | 2,900 (Bobcat NA) | Bobcat | SAME |
| T300 | 3,000 | dealer listings, spec database | SAME (on the edge) |
| T7X | 3,056 | one dealer listing | SAME (weak source) |
| T870 | 3,650 (Bobcat NA; LA 3,525) | Bobcat | SAME |
| T190 | 1,900 vs 2,399 | spec databases | CONFLICT |
| T200, T250, T320 | 2,000 / 2,500 / 3,100 | spec databases, Bobcat non-current page | ADDED |
| S185, S205, S590, S595, S62, S630, S64, S64-2 | as in the table | Bobcat regional, dealers, spec databases | SAME |
| S66, S66-2, S650, S76, S300, S330, S740, S770, S86 | as in the table | Bobcat regional, dealers, trade press | SAME |
| S530 | 1,916 vs 2,008 | Bobcat regional | CONFLICT |
| S550 | 1,750 to 2,041 by variant | Bobcat regional, dealers | CONFLICT |
| S570 | 1,950 SAE vs 2,081 ISO | Bobcat regional, dealers | CONFLICT |
| S86-2 | 3,500 vs 3,499 | Bobcat NA and metric | CONFLICT (edge) |

Mini excavators were not re-checked: their class follows from the model number by the tonnes rule, not from a published figure.

No class was changed; the macro (SizeClassMap) and the Python reference gained the T200, T250 and T320 rows.

## Models whose placement most needs a human check

1. **S86-2** (3,500 lb, exactly on the 3500&Over edge; Bobcat metric 1,587 kg = 3,499 lb, just under the edge). The base S86 is 3,400, so the two may land in different classes. Dealer sheets conflict.
2. **S64, S64-2, T64** (2,300 lb, exactly on the 2300 edge). T64-2 is confirmed 2300<2600 by the sample; check that S64 and S64-2 fall the same way in real data.
3. **T550** (1,995 vs 2,027 by model year). The sample says 0<2000.
4. **T770** (3,475 vs 3,567 international) and **S770** (3,350 vs 3,459). Both close to 3,500. T770 is confirmed 3000<3500 by the sample.
5. **S76, T76, T76-2** (2,900), **S740** (3,100), **S86** (3,400): within 100 lb of an edge; counterweight options push S76 and S86 over it.
6. **S650** (2,690), **T650** (2,570), **S66 / S66-2** (2,400 to 2,570), **T62** (2,150 to 2,260): near the 2,600 or 2,300 edges.
7. **T7X** (3,056 current vs 2,900 to 2,950 at launch).
8. **S530, S550, S570, S590** (1,916 to 2,100): cluster around 2,000; S550 has the worst conflicts (S530, S550 and S570 now have conflicting Bobcat figures across 2,000).
9. **T870** (3,650 / 3,525 / 3,325 by source).
10. **E40, E50, E50z, E30**: sit exactly on an inclusive upper edge. If the business ever buckets by actual weight, E40 moves to >04-05.
11. **E63, E85, E145, E165**: placed by the name/10 rule only, not confirmed by report data.
12. **Variant suffixes** (z, e, i, W): the macro upper-cases the MODEL_GROUP and looks it up as written; if Snowflake emits "E17Z" it must be in the map (it is), if it emits "E17" the base row applies.
13. **T190** (1,900 vs 2,399 by source): legacy, probably a variant split.
