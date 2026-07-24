/* ============================================================================
   PROJECT 2 - MEDICARE PART D PRESCRIBER SPEND ANALYSIS (TEXAS, 2022)
   Level 3 - Student Choice
   Sunanda Thompson

   Dataset: CMS Medicare Part D Prescribers by Provider and Drug (2022)
   https://data.cms.gov/provider-summary-by-type-of-service/medicare-part-d-prescribers/medicare-part-d-prescribers-by-provider-and-drug

   Filtered at download to: State = TX | Specialty in (Internal Medicine,
   Family Practice, Cardiology, Psychiatry) | Year = 2022

   Scope: 908,137 rows x 22 columns as extracted

   HYPOTHESIS: Specialty type and drug choice are the primary drivers of
   Medicare Part D drug costs in Texas. A small number of specialties and
   high-cost drugs account for the majority of total spending, and
   providers who prescribe more brand-name drugs have a higher cost per
   beneficiary than those who primarily prescribe generics.

   I followed the Data Analytics Workflow for this project: Frame, Extract,
   Wrangle/Prepare, Analyze, and Interpret all happen here in SQL.
   Communicate happens in my Excel workbook and slide deck, both built
   directly from the query results below. Every analysis query is
   commented with the exact Excel tab name it feeds, so the SQL and the
   workbook tie back to each other.
   ============================================================================ */


/* ============================================================================
   SECTION 1 - EXTRACT
   I'm importing the raw CSV into Postgres completely untouched, so I
   always have an unmodified copy to fall back on. Every column comes in
   as TEXT on purpose, the file mixes numbers with blank cells
   (CMS-suppressed values), and importing straight to NUMERIC would fail
   on the first blank. I cast everything deliberately in Section 2.
   ============================================================================ */

DROP TABLE IF EXISTS medicare_partd_raw;

-- Every column is TEXT here on purpose (see note above) - no casting yet
CREATE TABLE medicare_partd_raw (
    prscrbr_npi              TEXT,   -- provider's National Provider Identifier (unique ID)
    prscrbr_last_org_name    TEXT,   -- provider's last name
    prscrbr_first_name       TEXT,   -- provider's first name
    prscrbr_city              TEXT,  -- city where the provider practices
    prscrbr_state_abrvtn      TEXT,  -- state abbreviation (filtered to TX at download)
    prscrbr_state_fips        TEXT,  -- state FIPS code, kept for reference only
    prscrbr_type               TEXT, -- provider specialty (filtered to my 4 specialties)
    prscrbr_type_src           TEXT, -- source CMS used to assign the specialty
    brnd_name                  TEXT, -- brand name of the drug billed
    gnrc_name                  TEXT, -- generic name of the drug billed
    tot_clms                   TEXT, -- total claims for this provider-drug row
    tot_30day_fills            TEXT, -- claims standardized to 30-day fill equivalents
    tot_day_suply               TEXT,-- total days of medication supplied
    tot_drug_cst                TEXT,-- total drug cost paid — my main cost metric
    tot_benes                   TEXT,-- total unique patients; blank = CMS-suppressed (<11 patients)
    ge65_sprsn_flag              TEXT, -- reference flag, dropped later, not analytical
    ge65_tot_clms                TEXT, -- same metrics as above, restricted to age 65+ patients
    ge65_tot_30day_fills          TEXT,
    ge65_tot_drug_cst              TEXT,
    ge65_tot_day_suply              TEXT,
    ge65_bene_sprsn_flag              TEXT, -- reference flag, dropped later, not analytical
    ge65_tot_benes                     TEXT
);


-- Known source-file limitation (not something I can fix in SQL): per the
-- CMS data dictionary, provider-drug rows with fewer than 11 total claims
-- are excluded from this file entirely before I ever received it. This
-- means every total in my analysis slightly undercounts true low-volume
-- prescribing. Documented as Step 12 in my Data Handling Summary.



-- I imported the CSV using pgAdmin's Import/Export wizard (right-click the
-- table > Import/Export Data > Format: csv, Header: Yes, Delimiter: , Quote: ").

-- Confirmed row count immediately after import: 908,137
SELECT COUNT(*) AS row_count FROM medicare_partd_raw;

-- Sanity-checked that my extraction filters actually held: 1 state, 4 specialties.
SELECT
    COUNT(DISTINCT prscrbr_state_abrvtn) AS state_count,
    COUNT(DISTINCT prscrbr_type)          AS specialty_count
FROM medicare_partd_raw;


/* ============================================================================
   SECTION 2 - WRANGLE / PREPARE
   This section is the SQL version of my Excel "Data Handling Summary" tab.
   Every step here is documented there too.
   ============================================================================ */

-- Checked for duplicate provider-drug rows before doing any aggregation,
-- since duplicates would double-count claims and cost. Returned 0 rows.
SELECT prscrbr_npi, brnd_name, gnrc_name, COUNT(*) AS dup_count
FROM medicare_partd_raw
GROUP BY prscrbr_npi, brnd_name, gnrc_name
HAVING COUNT(*) > 1;   -- only shows rows that appear more than once

-- Counted how many rows CMS suppressed (blanked out) for privacy. This
-- happens automatically whenever a provider-drug combination serves fewer
-- than 11 patients. These are legitimate suppressions, not bad data, so I
-- do NOT fill them with 0. I keep them NULL and exclude them explicitly
-- wherever they'd affect a calculation.
-- Result: tot_benes_blank = 485,663 | ge65_clms_blank = 401,813 | ge65_benes_blank = 797,344
SELECT
    COUNT(*) FILTER (WHERE tot_benes = '')      AS tot_benes_blank,     -- counts only rows
    COUNT(*) FILTER (WHERE ge65_tot_clms = '')   AS ge65_clms_blank,    -- where that specific
    COUNT(*) FILTER (WHERE ge65_tot_benes = '')  AS ge65_benes_blank    -- column is blank text
FROM medicare_partd_raw;

-- Built my clean, analysis-ready table. This one statement handles every
-- remaining cleaning step at once: drops the two suppression-flag columns
-- (reference flags only, not analytical fields), casts every numeric
-- column from TEXT to NUMERIC using NULLIF(col, '') so blanks become true
-- NULLs instead of throwing a cast error, trims whitespace so GROUP BY
-- doesn't split "Houston " and "Houston" into separate buckets, and adds
-- 4 new columns I use throughout the analysis:
--   prscrbr_full_name - readable label for Q6 and Q10
--   drug_type         - 'Generic' when brand name = generic name, else
--                        'Brand'; this is what Q7 tests against my hypothesis
--   cost_per_claim / cost_per_bene - reusable unit-cost ratios
DROP TABLE IF EXISTS medicare_partd_clean;

CREATE TABLE medicare_partd_clean AS
SELECT
    prscrbr_npi,                                              -- kept as-is, already a clean ID

    TRIM(prscrbr_last_org_name)                               AS prscrbr_last_org_name,
    TRIM(prscrbr_first_name)                                  AS prscrbr_first_name,

    TRIM(prscrbr_first_name) || ' ' || TRIM(prscrbr_last_org_name)
                                                               -- concatenates first + last into
                                                               -- one readable label
                                                               AS prscrbr_full_name,

    TRIM(prscrbr_city)                                        AS prscrbr_city,
    prscrbr_state_abrvtn,                                     -- no cleaning needed, all 'TX'
    prscrbr_state_fips,                                       -- kept for reference only
    TRIM(prscrbr_type)                                        AS prscrbr_type,
    prscrbr_type_src,                                         -- no cleaning needed

    TRIM(brnd_name)                                           AS brnd_name,
    TRIM(gnrc_name)                                           AS gnrc_name,

    CASE WHEN LOWER(TRIM(brnd_name)) = LOWER(TRIM(gnrc_name)) -- case-insensitive compare: if the
         THEN 'Generic'                                       -- brand name and generic name are
         ELSE 'Brand' END                                     -- the same, it was billed under its
                                                               -- generic name; if they differ, a
                                                               AS drug_type,                   -- branded product was billed

    -- NULLIF(col, '') turns a blank string into a true SQL NULL BEFORE
    -- the ::NUMERIC cast runs, so an empty cell doesn't throw a cast error
    NULLIF(tot_clms, '')::NUMERIC                              AS tot_clms,
    NULLIF(tot_30day_fills, '')::NUMERIC                        AS tot_30day_fills,
    NULLIF(tot_day_suply, '')::NUMERIC                           AS tot_day_suply,
    NULLIF(tot_drug_cst, '')::NUMERIC                             AS tot_drug_cst,

    -- same blank-safe cast; stays NULL wherever CMS suppressed the count
    -- (<11 patients) - NOT zero-filled, since 0 would understate real patients
    NULLIF(tot_benes, '')::NUMERIC                                 AS tot_benes,

    -- same pattern, all restricted to the 65+ subgroup
    NULLIF(ge65_tot_clms, '')::NUMERIC                              AS ge65_tot_clms,
    NULLIF(ge65_tot_30day_fills, '')::NUMERIC                        AS ge65_tot_30day_fills,
    NULLIF(ge65_tot_drug_cst, '')::NUMERIC                            AS ge65_tot_drug_cst,
    NULLIF(ge65_tot_day_suply, '')::NUMERIC                            AS ge65_tot_day_suply,

    -- suppressed even more often than the all-ages tot_benes column
    NULLIF(ge65_tot_benes, '')::NUMERIC                                 AS ge65_tot_benes,

    -- total cost ÷ total claims; the inner NULLIF(..., 0) turns a 0 claim
    -- count into NULL so I never get a divide-by-zero error
    ROUND(NULLIF(tot_drug_cst, '')::NUMERIC
          / NULLIF(NULLIF(tot_clms, '')::NUMERIC, 0), 2)        AS cost_per_claim,

    -- total cost ÷ total beneficiaries; stays NULL wherever tot_benes is
    -- suppressed - expected, not a bug
    ROUND(NULLIF(tot_drug_cst, '')::NUMERIC
          / NULLIF(NULLIF(tot_benes, '')::NUMERIC, 0), 2)         AS cost_per_bene

FROM medicare_partd_raw;   -- source: the untouched raw import

-- Confirmed the clean table landed correctly: 908,137 rows.
SELECT COUNT(*) AS clean_row_count FROM medicare_partd_clean;

-- Re-ran the null count against the clean table to make sure nothing got
-- accidentally zero-filled during casting. Matched exactly: 485,663 / 401,813 / 797,344.
SELECT
    COUNT(*) FILTER (WHERE tot_benes IS NULL)      AS tot_benes_null,
    COUNT(*) FILTER (WHERE ge65_tot_clms IS NULL)  AS ge65_clms_null,
    COUNT(*) FILTER (WHERE ge65_tot_benes IS NULL) AS ge65_benes_null
FROM medicare_partd_clean;

-- Indexed the columns I group and join on repeatedly across the next 10
-- queries (specialty, provider ID, city), so Postgres can jump straight
-- to matching rows instead of scanning all 908K rows every time.
CREATE INDEX idx_clean_type ON medicare_partd_clean (prscrbr_type);  -- used by Q1, Q3, Q4, Q10
CREATE INDEX idx_clean_npi  ON medicare_partd_clean (prscrbr_npi);   -- used by Q5, Q6, Q7, Q10
CREATE INDEX idx_clean_city ON medicare_partd_clean (prscrbr_city);  -- used by Q5


/* ============================================================================
   SECTION 3 - REFERENCE TABLE (supports the JOIN used in Q5)
   Before I could build a city -> region lookup table, I needed to know
   which cities actually show up as the top spenders, so I ran the city
   ranking first, then built the reference table from those results.
   ============================================================================ */

-- Preliminary look at the top 10 TX cities by total drug cost.
SELECT
    prscrbr_city,
    COUNT(DISTINCT prscrbr_npi) AS total_prescribers,   -- unique providers in this city
    SUM(tot_drug_cst)           AS total_city_cost      -- total drug spend for this city
FROM medicare_partd_clean
GROUP BY prscrbr_city
HAVING COUNT(DISTINCT prscrbr_npi) >= 5   -- dropped tiny cities so the ranking isn't skewed
ORDER BY total_city_cost DESC
LIMIT 10;

-- Result: Houston, San Antonio, Dallas, El Paso, Austin, Fort Worth,
-- Corpus Christi, Lubbock, Mcallen, Plano. I built my region lookup table
-- from these 10 cities.
DROP TABLE IF EXISTS tx_city_region;

CREATE TABLE tx_city_region (
    city    TEXT PRIMARY KEY,   -- each city can only appear once
    region  TEXT NOT NULL       -- every city must have a region assigned
);

INSERT INTO tx_city_region (city, region) VALUES
    ('Houston',        'Gulf Coast'),
    ('San Antonio',    'South Texas'),
    ('Dallas',         'North Texas'),
    ('El Paso',        'West Texas'),
    ('Austin',         'Central Texas'),
    ('Fort Worth',     'North Texas'),
    ('Corpus Christi', 'Coastal Bend'),
    ('Lubbock',        'South Plains'),
    ('Mcallen',        'Rio Grande Valley'),
    ('Plano',          'North Texas');

-- Confirmed all 10 rows landed correctly.
SELECT * FROM tx_city_region;


/* ============================================================================
   SECTION 4 - ANALYZE
   Ten business questions, each commented with the exact Excel tab it
   feeds, the fields combined, and the result I confirmed when I ran it.
   ============================================================================ */

-- -----------------------------------------------------------------------
-- 1. spend_by_specialty
-- Which of the 4 specialties accounts for the highest total Part D drug
-- cost, and what % of total spend does each represent?
--
-- SUM(...) OVER () gives me a grand total across all specialties without
-- collapsing the GROUP BY rows, so I can calculate % of total in the same
-- query instead of running a second one just for the denominator.
--
-- Result: Family Practice $2.00B (41.67%) | Internal Medicine $1.90B
-- (39.49%) | Cardiology $687M (14.30%) | Psychiatry $218M (4.54%)
-- -----------------------------------------------------------------------
SELECT
    prscrbr_type,
    ROUND(SUM(tot_drug_cst), 2)                         AS total_drug_cost,   -- this specialty's total

    ROUND(
        100.0 * SUM(tot_drug_cst)                       -- this specialty's total cost ...
        / SUM(SUM(tot_drug_cst)) OVER ()                -- ... ÷ the grand total across ALL
                                                         -- specialties (the window function)
    , 2)                                                 AS pct_of_total_spend

FROM medicare_partd_clean
GROUP BY prscrbr_type
ORDER BY total_drug_cost DESC;


-- -----------------------------------------------------------------------
-- 2. top_generic_drugs
-- Which are the top 10 most prescribed generic drugs by total claim
-- count, and how does their cost rank compare to their claim-count rank?
--
-- I used a CTE to stage this in two steps: first total up claims/cost per
-- drug, then rank those totals two different ways (RANK() by claims, and
-- separately by cost) so I can compare the two side by side.
--
-- Result (top row): Atorvastatin Calcium - claim_rank 1, cost_rank 23
-- ($37.5M from 2,940,737 claims)
-- -----------------------------------------------------------------------
WITH drug_totals AS (
    -- STAGE 1: total claims and cost per drug
    SELECT
        gnrc_name,
        SUM(tot_clms)      AS total_claims,
        SUM(tot_drug_cst)  AS total_drug_cost
    FROM medicare_partd_clean
    GROUP BY gnrc_name
),
drug_ranks AS (
    -- STAGE 2: reads Stage 1 like a table, assigns two independent rankings
    SELECT
        gnrc_name,
        total_claims,
        ROUND(total_drug_cost, 2)                    AS total_drug_cost,
        RANK() OVER (ORDER BY total_claims DESC)     AS claim_rank,   -- #1 = most claims of any drug
        RANK() OVER (ORDER BY total_drug_cost DESC)  AS cost_rank     -- #1 = most expensive drug
    FROM drug_totals
)
SELECT *
FROM drug_ranks
ORDER BY claim_rank
LIMIT 10;


-- -----------------------------------------------------------------------
-- 3. cost_per_claim_specialty
-- What's the average drug cost per claim by specialty, and which
-- specialty shows the largest gap between its claim-count share and its
-- cost share?
--
-- Positive gap = disproportionately expensive; negative gap
-- = high-volume but cheap.
--
-- Result: Cardiology highest avg cost/claim ($152.65) and largest
-- positive gap (+5.47 pts). Family Practice largest negative gap (-8.51 pts)
-- -----------------------------------------------------------------------
SELECT
    prscrbr_type,
    SUM(tot_clms)                                                       AS total_claims,
    ROUND(SUM(tot_drug_cst), 2)                                          AS total_drug_cost,

    ROUND(SUM(tot_drug_cst) / SUM(tot_clms), 2)                           -- total cost ÷ total claims
                                                                         AS avg_cost_per_claim,

    ROUND(100.0 * SUM(tot_clms) / SUM(SUM(tot_clms)) OVER (), 2)          -- this specialty's % of
                                                                         AS pct_of_total_claims,     -- all claims

    ROUND(100.0 * SUM(tot_drug_cst) / SUM(SUM(tot_drug_cst)) OVER (), 2)  -- this specialty's % of
                                                                         AS pct_of_total_cost,        -- all cost

    ROUND(
        (100.0 * SUM(tot_drug_cst) / SUM(SUM(tot_drug_cst)) OVER ())     -- cost share ...
      - (100.0 * SUM(tot_clms)     / SUM(SUM(tot_clms))     OVER ()), 2) -- ... minus claim share
                                                                         AS cost_share_minus_claim_share

FROM medicare_partd_clean
GROUP BY prscrbr_type
ORDER BY avg_cost_per_claim DESC;


-- -----------------------------------------------------------------------
-- 4. cost_per_bene_specialty
-- Which specialty has the highest/lowest cost per beneficiary in Texas?
--
-- I filtered out CMS-suppressed rows.
--
-- Result: Cardiology highest ($441.98/beneficiary), Family Practice
-- lowest ($150.67/beneficiary)
-- -----------------------------------------------------------------------
SELECT
    prscrbr_type,
    ROUND(SUM(tot_drug_cst), 2)                    AS total_drug_cost,
    SUM(tot_benes)                                  AS total_beneficiaries,
    ROUND(SUM(tot_drug_cst) / SUM(tot_benes), 2)    AS cost_per_beneficiary   -- cost ÷ beneficiaries
FROM medicare_partd_clean
WHERE tot_benes IS NOT NULL   -- excludes CMS-suppressed rows (<11 patients) — SUM() would skip
                              -- NULLs anyway, but this makes the exclusion explicit and readable
GROUP BY prscrbr_type
ORDER BY cost_per_beneficiary DESC;


-- -----------------------------------------------------------------------
-- 5. high_cost_cities
-- Which TX cities have the highest concentration of high-cost
-- prescribers, and does location influence total drug spend?
--
-- Three-stage CTE: (1) roll cost up to one row per provider, (2) find the
-- 75th percentile of provider cost using PERCENTILE_CONT - that's my
-- definition of "high-cost" - then CROSS JOIN that single threshold value
-- onto every provider row so I can compare each one against it, (3) group
-- by city and count how many providers cleared the threshold. I finish
-- with an INNER JOIN to my tx_city_region lookup table to pull in
-- each city's region.
--
-- Result (top row): Houston - 2,134 prescribers, $602.7M, 20.48%
-- high-cost, Gulf Coast. McAllen has the highest concentration at 38.62%.
-- -----------------------------------------------------------------------
WITH provider_totals AS (
    -- STAGE 1: one row per provider (a provider has many rows in
    -- medicare_partd_clean, one per drug, so I collapse that down first)
    SELECT
        prscrbr_npi,
        MIN(prscrbr_city)   AS city,          -- same city on every row for a given provider,
                                               -- MIN just picks that one value
        SUM(tot_drug_cst)   AS provider_total_cost
    FROM medicare_partd_clean
    GROUP BY prscrbr_npi
),
threshold AS (
    -- STAGE 2: the ONE dollar value that defines "high-cost" — the 75th percentile
    SELECT PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY provider_total_cost) AS p75
    FROM provider_totals
),
city_summary AS (
    -- STAGE 3: group providers by city, count how many are above the threshold
    SELECT
        pt.city,
        COUNT(*)                                                        AS total_prescribers,
        COUNT(*) FILTER (WHERE pt.provider_total_cost > t.p75)           AS high_cost_prescribers,
        SUM(pt.provider_total_cost)                                     AS total_city_cost
    FROM provider_totals pt
    CROSS JOIN threshold t   -- attaches the single p75 value onto every provider row
    GROUP BY pt.city
    HAVING COUNT(*) >= 5     -- drops tiny cities so the % metric isn't skewed
)
SELECT
    cs.city,
    cs.total_prescribers,
    cs.high_cost_prescribers,
    ROUND(cs.total_city_cost, 2)                                         AS total_city_cost,
    ROUND(100.0 * cs.high_cost_prescribers / cs.total_prescribers, 2)     AS pct_high_cost,
    r.region                                                              -- pulled in from the JOIN
FROM city_summary cs
INNER JOIN tx_city_region r                -- the real two-table JOIN
    ON cs.city = r.city                    -- matches each city to its region
ORDER BY cs.total_city_cost DESC
LIMIT 10;


-- -----------------------------------------------------------------------
-- 6. top_20_prescribers
-- Who are the top 20 individual prescribers by total drug cost, and what
-- specialties do they represent?
--
-- Result (top row): Nina Olvera, Family Practice, San Antonio -
-- $23,709,303.17 from just 4,531 claims
-- -----------------------------------------------------------------------
SELECT
    prscrbr_npi,
    prscrbr_full_name,
    prscrbr_type,
    prscrbr_city,
    ROUND(SUM(tot_drug_cst), 2)  AS total_drug_cost,
    SUM(tot_clms)                AS total_claims
FROM medicare_partd_clean
GROUP BY prscrbr_npi, prscrbr_full_name, prscrbr_type, prscrbr_city   -- Postgres requires every
                                                                       -- non-aggregated SELECT column
                                                                       -- to appear in GROUP BY; since
                                                                       -- name/type/city are all tied
                                                                       -- 1-to-1 to a single npi, this
                                                                       -- doesn't change the result
ORDER BY total_drug_cost DESC
LIMIT 20;


-- -----------------------------------------------------------------------
-- 7. brand_vs_generic
-- Do providers who primarily prescribe brand-name drugs spend more per
-- patient than providers who primarily prescribe generics?
--
-- This is the query that directly tests my hypothesis.
--
-- Three-stage CTE: (1) split each provider's claims into brand vs generic
-- totals using FILTER, (2) classify each provider "Primarily Brand" if
-- over 50% of their claims are brand-name (COALESCE guards against a
-- provider who prescribed zero brand or zero generic claims, so I'm never
-- dividing by a NULL), (3) separately roll up each provider's total cost
-- and patient count. I then JOIN the classification to the cost data and
-- average cost-per-patient within each group.
--
-- Result: Primarily Brand - 496 providers, $1,614.73/patient. Primarily
-- Generic - 13,478 providers, $151.83/patient. Roughly a 10.6x gap -
-- my strongest evidence for the hypothesis.
-- -----------------------------------------------------------------------
WITH provider_claim_mix AS (
    -- STAGE 1: FILTER sums only the rows matching each condition, so I get
    -- brand claims and generic claims as two separate columns in one pass
    SELECT
        prscrbr_npi,
        SUM(tot_clms) FILTER (WHERE drug_type = 'Brand')   AS brand_claims,
        SUM(tot_clms) FILTER (WHERE drug_type = 'Generic') AS generic_claims
    FROM medicare_partd_clean
    GROUP BY prscrbr_npi
),
provider_group AS (
    -- STAGE 2: classify each provider based on their brand %
    SELECT
        prscrbr_npi,
        CASE
            WHEN COALESCE(brand_claims, 0)::NUMERIC                          -- COALESCE swaps a
                 / NULLIF(COALESCE(brand_claims, 0)                          -- NULL for 0 (a provider
                          + COALESCE(generic_claims, 0), 0) > 0.5             -- with zero brand or zero
            THEN 'Primarily Brand'                                           -- generic claims), so the
            ELSE 'Primarily Generic'                                         -- division never breaks;
        END AS provider_group                                                -- NULLIF(...,0) on the
    FROM provider_claim_mix                                                  -- denominator guards
                                                                               -- against divide-by-zero
),
provider_cost AS (
    -- STAGE 3: separately, each provider's total cost and total patients
    SELECT
        prscrbr_npi,
        SUM(tot_drug_cst) AS total_cost,
        SUM(tot_benes)    AS total_benes
    FROM medicare_partd_clean
    WHERE tot_benes IS NOT NULL   -- excludes CMS-suppressed rows
    GROUP BY prscrbr_npi
)
SELECT
    pg.provider_group,
    COUNT(*)                                                                AS provider_count,
    ROUND(AVG(pc.total_cost / pc.total_benes), 2)                           AS avg_cost_per_patient
FROM provider_group pg
INNER JOIN provider_cost pc
    ON pg.prscrbr_npi = pc.prscrbr_npi   -- matches each provider's label to their cost numbers
GROUP BY pg.provider_group
ORDER BY avg_cost_per_patient DESC;


-- -----------------------------------------------------------------------
-- 8. drug_patient_reach
-- Which drugs reach the largest number of unique patients, and which get
-- prescribed most repeatedly to the same patients?
--
-- Two separate queries off the same base table,
-- sorted differently to answer each half.
--
-- 8a - widest reach (sorted by total patients)
--
-- Result (top row): Atorvastatin Calcium - 930,278 patients
-- -----------------------------------------------------------------------
SELECT
    gnrc_name,
    SUM(tot_benes)                                 AS total_benes,
    SUM(tot_clms)                                  AS total_claims,
    ROUND(SUM(tot_clms) / SUM(tot_benes), 2)        AS claims_per_bene   -- avg refills per patient
FROM medicare_partd_clean
WHERE tot_benes IS NOT NULL
GROUP BY gnrc_name
ORDER BY total_benes DESC     -- sorted by REACH: most unique patients first
LIMIT 10;

-- -----------------------------------------------------------------------
-- 8b - most repeated prescribing per patient (chronic refill pattern).
--
-- I added a HAVING filter (100+ total patients) to drop drugs with a tiny
-- patient pool, which would otherwise produce misleadingly extreme ratios.
--
-- Result (top row): Clozapine - 11.08 claims per beneficiary
-- (356 patients, 3,945 claims)
-- -----------------------------------------------------------------------
SELECT
    gnrc_name,
    SUM(tot_benes)                                 AS total_benes,
    SUM(tot_clms)                                  AS total_claims,
    ROUND(SUM(tot_clms) / SUM(tot_benes), 2)        AS claims_per_bene
FROM medicare_partd_clean
WHERE tot_benes IS NOT NULL
GROUP BY gnrc_name
HAVING SUM(tot_benes) >= 100   -- drops small-sample drugs that would otherwise produce a
                                -- misleadingly extreme ratio (e.g. 2 patients, 20 claims)
ORDER BY claims_per_bene DESC   -- sorted by RATIO: most refills per patient first
LIMIT 10;


-- -----------------------------------------------------------------------
-- 9. ge65_vs_all
-- How does prescribing volume and cost for beneficiaries 65+ compare to
-- the total population?
--
-- I'm comparing four different metrics as rows rather than grouping
-- by a column that already exists in the data, so I
-- wrote four independent SELECTs and stacked them with UNION ALL (which
-- keeps every row, unlike plain UNION which would also try to de-duplicate).
--
-- Note: the beneficiary row is suppressed far more heavily in the 65+
-- column than the all-ages column, so that % understates the true 65+
-- share - I don't compare it directly to the claims/cost percentages.
--
-- Result: Claims 56.24% are 65+ | Cost 51.47% are 65+ | Beneficiaries
-- 28.37% (understated, see note) | Avg cost/claim ratio 91.52% (65+
-- patients actually cost LESS per claim, not more)
-- -----------------------------------------------------------------------

-- BLOCK 1: claims comparison
SELECT
    'Total Claims'      AS metric,                            -- a literal label, not a table
                                                                -- column — becomes this row's name
    SUM(tot_clms)        AS all_ages,
    SUM(ge65_tot_clms)    AS age_65_plus,
    ROUND(100.0 * SUM(ge65_tot_clms) / SUM(tot_clms), 2) AS pct_65plus_of_all
FROM medicare_partd_clean

UNION ALL   -- stacks the next block as additional rows underneath this one

-- BLOCK 2: cost comparison, same pattern as Block 1
SELECT
    'Total Drug Cost',
    ROUND(SUM(tot_drug_cst), 2),
    ROUND(SUM(ge65_tot_drug_cst), 2),
    ROUND(100.0 * SUM(ge65_tot_drug_cst) / SUM(tot_drug_cst), 2)
FROM medicare_partd_clean

UNION ALL

-- BLOCK 3: beneficiary comparison — FILTER drops suppressed rows from
-- both the all-ages and 65+ sums before they're compared
SELECT
    'Total Beneficiaries (non-suppressed, not de-duplicated)',
    SUM(tot_benes) FILTER (WHERE tot_benes IS NOT NULL),
    SUM(ge65_tot_benes) FILTER (WHERE ge65_tot_benes IS NOT NULL),   -- suppressed more heavily
                                                                     -- than tot_benes - see note
    ROUND(100.0 * SUM(ge65_tot_benes) FILTER (WHERE ge65_tot_benes IS NOT NULL)
          / SUM(tot_benes) FILTER (WHERE tot_benes IS NOT NULL), 2)
FROM medicare_partd_clean

UNION ALL

-- BLOCK 4: avg cost per claim — a ratio-of-ratios, computed fresh for
-- each population rather than reusing Block 1/2's totals
SELECT
    'Avg Cost per Claim',
    ROUND(SUM(tot_drug_cst) / SUM(tot_clms), 2),
    ROUND(SUM(ge65_tot_drug_cst) / SUM(ge65_tot_clms), 2),
    ROUND(100.0 * (SUM(ge65_tot_drug_cst) / SUM(ge65_tot_clms))     -- if this is under 100%,
                 / (SUM(tot_drug_cst) / SUM(tot_clms)), 2)           -- 65+ costs LESS per claim
FROM medicare_partd_clean;


-- -----------------------------------------------------------------------
-- 10. provider_cost_outliers
-- Are there individual providers whose total cost is significantly higher
-- than other providers in the same specialty, worth flagging for review?
--
-- I used a window function with PARTITION BY to get each specialty's
-- average and standard deviation attached to every provider row (no
-- GROUP BY collapsing needed), then calculated a z-score - how many
-- standard deviations above their own specialty's average each provider
-- sits - and kept anyone above 3 (statistically, roughly the top 0.1%).
--
-- Result: 220 providers actually clear the z-score > 3 threshold across
-- all 4 specialties (107 Internal Medicine, 77 Family Practice,
-- 19 Psychiatry, 17 Cardiology). The LIMIT 15 below caps the output at
-- the 15 most extreme by z-score, not the full outlier population.
-- This is a "top 15" list, not "the 15 outliers that exist."
-- Top row: Nina Olvera, Family Practice, $23,709,303.17, z-score ~40.9
-- -----------------------------------------------------------------------
WITH provider_totals AS (
    -- STAGE 1: one row per provider (same pattern as Q5/Q6)
    SELECT
        prscrbr_npi,
        MAX(prscrbr_full_name) AS prscrbr_full_name,   -- same name on every row for a
                                                         -- given provider, MAX just picks it
        prscrbr_type,
        SUM(tot_drug_cst)      AS total_cost
    FROM medicare_partd_clean
    GROUP BY prscrbr_npi, prscrbr_type
),
provider_zscores AS (
    -- STAGE 2: PARTITION BY computes a SEPARATE average/stddev for each
    -- specialty (not one grand total like Q1/Q3's OVER ()), and attaches
    -- it to every provider row in that specialty — no GROUP BY collapsing
    SELECT
        prscrbr_npi,
        prscrbr_full_name,
        prscrbr_type,
        total_cost,
        AVG(total_cost)    OVER (PARTITION BY prscrbr_type) AS specialty_avg_cost,
        STDDEV(total_cost) OVER (PARTITION BY prscrbr_type) AS specialty_stddev_cost
    FROM provider_totals
)
SELECT
    prscrbr_npi,
    prscrbr_full_name,
    prscrbr_type,
    ROUND(total_cost, 2)                                                    AS total_drug_cost,
    ROUND(specialty_avg_cost, 2)                                             AS specialty_avg_cost,
    ROUND(specialty_stddev_cost, 2)                                          AS specialty_stddev_cost,
    ROUND((total_cost - specialty_avg_cost) / specialty_stddev_cost, 1)      AS z_score   -- z-score
                                                                                            -- formula:
                                                                                            -- (value -
                                                                                            -- average)
                                                                                            -- ÷ stddev
FROM provider_zscores
WHERE (total_cost - specialty_avg_cost) / specialty_stddev_cost > 3     -- keep only providers
                                                                        -- more than 3 std
                                                                        -- deviations above
                                                                        -- their specialty's average
                                                                        -- (220 providers meet
                                                                        -- this condition)
ORDER BY z_score DESC
LIMIT 15;   -- caps the result at the 15 MOST EXTREME of those 220,
            -- not the total count of outliers




-- -----------------------------------------------------------------------
-- 10b. provider_cost_distribution
-- Supports a box-and-whisker chart (or scatter-plus-threshold-line
-- chart) that visually proves the outliers in Q10 are real, instead of
-- just asserting it with a z-score number. 
--
-- Note: Q10's LIMIT 15 only
-- shows the 15 MOST EXTREME providers. 220 total actually clear the
-- z-score > 3 threshold. This distribution query (with no WHERE filter
-- or LIMIT) is what makes that visible: the box-plot dots above each
-- specialty's box will show all 220, with the top 15 highlighted
-- separately in the chart itself.
--
-- This is almost the same query as Q10, but with two changes: the WHERE
-- filter that kept only z-score > 3 is removed, and the LIMIT 15 is
-- removed. A box-and-whisker chart needs the FULL distribution of every
-- provider's cost, grouped by specialty, to draw the "normal range" box
-- in the first place. You can't show something is an outlier without
-- also showing what "not an outlier" looks like for that specialty.
--
-- I kept the z_score column in the output even though the chart doesn't
-- strictly need it, since it's useful for conditional formatting or for
-- building the scatter-plus-threshold-line alternative chart later.
--
-- Result: 15,863 providers, all 4 specialties, sorted by specialty so
-- Excel can group cleanly when building the chart.
-- -----------------------------------------------------------------------
WITH provider_totals AS (
    -- STAGE 1: one row per provider (identical to Q10, Stage 1)
    SELECT
        prscrbr_npi,
        MAX(prscrbr_full_name) AS prscrbr_full_name,   -- same name on every row for a
                                                         -- given provider, MAX just picks it
        prscrbr_type,
        SUM(tot_drug_cst)      AS total_cost
    FROM medicare_partd_clean
    GROUP BY prscrbr_npi, prscrbr_type
),
provider_zscores AS (
    -- STAGE 2: identical to Q10, Stage 2 - PARTITION BY still computes a
    -- SEPARATE average/stddev for each specialty, attached to every row
    SELECT
        prscrbr_npi,
        prscrbr_full_name,
        prscrbr_type,
        total_cost,
        AVG(total_cost)    OVER (PARTITION BY prscrbr_type) AS specialty_avg_cost,
        STDDEV(total_cost) OVER (PARTITION BY prscrbr_type) AS specialty_stddev_cost
    FROM provider_totals
)
SELECT
    prscrbr_npi,
    prscrbr_full_name,
    prscrbr_type,
    ROUND(total_cost, 2)                                                    AS total_drug_cost,
    ROUND(specialty_avg_cost, 2)                                             AS specialty_avg_cost,
    ROUND(specialty_stddev_cost, 2)                                          AS specialty_stddev_cost,
    ROUND((total_cost - specialty_avg_cost) / specialty_stddev_cost, 1)      AS z_score   -- same
                                                                                            -- formula
                                                                                            -- as Q10:
                                                                                            -- (value -
                                                                                            -- average)
                                                                                            -- ÷ stddev
FROM provider_zscores
-- NOTE: no WHERE filter here on purpose. Q10 filtered to z-score > 3
-- (220 providers) then capped the output at the top 15 most extreme;
-- this version keeps every provider, including all 220 that clear the
-- threshold, since the chart needs the whole distribution to draw a
-- meaningful box per specialty.
ORDER BY prscrbr_type, total_drug_cost DESC;
-- NOTE: no LIMIT here either, for the same reason. This will return
-- all ~15,863 providers, which is expected and needed for this chart.




-- -----------------------------------------------------------------------
-- REFERENCE - day_supply_per_claim_specialty 
-- Exploratory query: does prescription length (day supply) vary meaningfully 
-- by specialty? Pulled during initial data exploration but not promoted to 
-- a full business question, since it did not tie back to cost or patient 
-- impact the way the other 10 questions did.
-- -----------------------------------------------------------------------
SELECT
    prscrbr_type,
    ROUND(SUM(tot_drug_cst), 2)                         AS total_drug_cost,   -- this specialty's total

    ROUND(
        100.0 * SUM(tot_drug_cst)                       -- this specialty's total cost ...
        / SUM(SUM(tot_drug_cst)) OVER ()                -- ... ÷ the grand total across ALL
                                                         -- specialties (the window function)
    , 2)                                                 AS pct_of_total_spend

FROM medicare_partd_clean
GROUP BY prscrbr_type
ORDER BY total_drug_cost DESC;
