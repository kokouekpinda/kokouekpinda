/*******************************************************************************
*  18_Strand_Report_Additions.do
*
*  Purpose : Reproduce every new table proposed for the qualitative and
*            quantitative strand reports of the REALISE final evaluation
*            (proposals document of 4 November 2026).
*  Reads   : $REALISE_DATA_FOLDER/constructed_<ssb|lipw|clas|sct>.dta
*  Writes  : $REALISE_OUTPUT_FOLDER/18_Strand_Report_Additions.xlsx
*            one sheet per table, plus a reproduction check of P1 to P4
*  Design  : survey design of the primary analysis: final response weight,
*            exact sampling-cell strata, respondent PSUs, finite-population
*            correction N_h, singleton strata centred. Proportions use a
*            logit-transformed 95% interval; scores use a symmetric interval.
*  Domains : a domain estimate keeps every interview in the design; cases
*            outside the domain or without a valid answer enter with zero
*            weight (subpopulation estimation).
*  Rules   : valid n below 10 suppressed; 10 to 24 flagged indicative.
*            Column J flags 0/1 results whose positive or negative count is
*            1 to 9 (complementary-cell precaution of the quantitative report).
*            Weighted estimates; n unweighted. Descriptive only. Tables not
*            listed in the PAP are exploratory (marked in the sheet title).
*  Author  : Romaric Ekpinda
*  Version : 4 November 2026. Requires Stata 17.
*******************************************************************************/

version 17.0
set more off
set varabbrev off


/*******************************************************************************
*  SECTION 1. FOLDERS AND OUTPUT WORKBOOK
*******************************************************************************/

global REALISE_DATA_FOLDER   "<root>/data/constructed"
global REALISE_OUTPUT_FOLDER "<root>/outputs/strand_additions"
global REALISE_OUTPUT_WORKBOOK "$REALISE_OUTPUT_FOLDER/18_Strand_Report_Additions.xlsx"

capture erase "$REALISE_OUTPUT_WORKBOOK"

**** Small-cell rules: suppressed below 10 valid answers, indicative from 10 to 24
local minimum_reported_n 10
local minimum_firm_n     25


/*******************************************************************************
*  SECTION 2. STACKED ANALYSIS FILE WITH HARMONISED INDICATORS
*  One record per interview, four components appended. Indicators are 0/1
*  among valid answers and missing otherwise; codes 98, 99 and the extended
*  missing codes are excluded from valid answers.
*******************************************************************************/

tempfile stacked_analysis_file
local component_list "ssb lipw clas sct"
local first_component "ssb"

foreach component of local component_list {

    use "$REALISE_DATA_FOLDER/constructed_`component'.dta", clear

    **** CLAS-only items are created as missing in the other components
    foreach clas_item in k12a k22 k24 {
        capture confirm variable `clas_item'
        if _rc generate `clas_item' = .
    }

    keep psu_id sampling_cell_id N_h final_weight sex_bin round_num county_code ///
         p1 p2 p3 p4 p2_noschedule g01 g03 g05 g05b g06 g07 r06b k12a k22 k24

    generate str4 component_name = "`component'"

    **** Overall satisfaction (G07): satisfied or very satisfied
    generate satisfied_overall = inlist(g07, 4, 5) if inlist(g07, 1, 2, 3, 4, 5)

    **** Reporting safety (G06): safe or very safe; unsafe or very unsafe
    generate safe_to_report_misconduct   = inlist(g06, 4, 5) if inlist(g06, 1, 2, 3, 4, 5)
    generate unsafe_to_report_misconduct = inlist(g06, 1, 2) if inlist(g06, 1, 2, 3, 4, 5)

    **** Complaint pathway: knowledge (G01), use (G03), access when needed (G05B), resolution (G05)
    generate knows_complaint_channel   = (g01 == 1) if inlist(g01, 0, 1)
    generate used_complaint_channel    = (g03 == 1) if inlist(g03, 0, 1)
    generate channel_access_easy       = (g05b == 1) if inlist(g05b, 1, 2, 3)
    generate channel_not_accessible    = (g05b == 3) if inlist(g05b, 1, 2, 3)
    generate complaint_resolved_any    = inlist(g05, 1, 2) if inlist(g05, 0, 1, 2)

    **** Delivery and change indicators used in the county and round tables
    generate no_schedule_communicated  = p2_noschedule
    generate worse_off_basic_needs     = (p4 < 0) if !missing(p4)

    **** Own or joint final decision on the use of support (R06B); codes differ by questionnaire
    if "`component'" == "sct" {
        generate own_or_joint_decision = inlist(r06b, 1, 2, 4) if !missing(r06b) & !inlist(r06b, 98, 99)
    }
    else {
        generate own_or_joint_decision = inlist(r06b, 1, 3) if !missing(r06b) & !inlist(r06b, 98, 99)
    }

    **** CLAS input package received in full (K12A) and first harvest completed (K22)
    generate clas_full_input_package   = (k12a == 1) if inlist(k12a, 0, 1, 2)
    generate clas_harvest_completed    = (k22 == 1) if inlist(k22, 0, 1)

    **** CLAS group expected to continue farming after the project: likely or very likely (K24)
    generate clas_continuation_likely  = inlist(k24, 4, 5) if inlist(k24, 1, 2, 3, 4, 5)

    if "`component'" != "`first_component'" append using "`stacked_analysis_file'"
    save "`stacked_analysis_file'", replace
}

**** Design identifiers unique across the stacked file
egen design_stratum = group(sampling_cell_id)
egen design_psu     = group(component_name psu_id)
label define sex_label 1 "Women" 2 "Men", replace
label values sex_bin sex_label
save "`stacked_analysis_file'", replace


/*******************************************************************************
*  SECTION 3. REPRODUCTION CHECK OF THE PUBLISHED PRIMARY OUTCOMES
*  Published values (31 October 2026 results tables): P1 and P2 in percent,
*  P3 and P4 in score points.
*******************************************************************************/

local published_ssb_p1  56.9
local published_ssb_p2  91.5
local published_ssb_p3  4.05
local published_ssb_p4  0.65
local published_lipw_p1 50.6
local published_lipw_p2 72.8
local published_lipw_p3 5.91
local published_lipw_p4 0.47
local published_clas_p1 32.7
local published_clas_p2 68.9
local published_clas_p3 5.35
local published_clas_p4 0.06
local published_sct_p1  61.1
local published_sct_p2  62.8
local published_sct_p3  6.53
local published_sct_p4  1.06

putexcel set "$REALISE_OUTPUT_WORKBOOK", sheet("Check_P1_P4") replace
putexcel A1 = "Reproduction check of published primary outcomes P1 to P4"
putexcel A2 = "Component" B2 = "Outcome" C2 = "Computed" D2 = "Published" ///
         E2 = "95% CI lower" F2 = "95% CI upper" G2 = "Valid n" H2 = "Status"
local check_row = 3

foreach component of local component_list {
    use "`stacked_analysis_file'", clear
    keep if component_name == "`component'"
    svyset design_psu [pweight = final_weight], strata(design_stratum) fpc(N_h) singleunit(centered)

    foreach outcome in p1 p2 p3 p4 {
        quietly svy: mean `outcome'
        matrix outcome_estimate = r(table)
        local point_estimate = outcome_estimate[1, 1]
        local standard_error = outcome_estimate[2, 1]
        local t_critical     = invttail(e(df_r), 0.025)
        quietly count if !missing(`outcome')
        local valid_n = r(N)

        if inlist("`outcome'", "p1", "p2") {
            local logit_se    = `standard_error' / (`point_estimate' * (1 - `point_estimate'))
            local lower_bound = 100 * invlogit(logit(`point_estimate') - `t_critical' * `logit_se')
            local upper_bound = 100 * invlogit(logit(`point_estimate') + `t_critical' * `logit_se')
            local computed_value = round(100 * `point_estimate', 0.1)
            local tolerance = 0.051
        }
        else {
            local lower_bound = `point_estimate' - `t_critical' * `standard_error'
            local upper_bound = `point_estimate' + `t_critical' * `standard_error'
            local computed_value = round(`point_estimate', 0.01)
            local tolerance = 0.0051
        }
        local check_status = cond(abs(`computed_value' - `published_`component'_`outcome'') < `tolerance', "Reproduced", "Differs")

        putexcel A`check_row' = "`component'" B`check_row' = "`outcome'" ///
                 C`check_row' = `computed_value' D`check_row' = `published_`component'_`outcome'' ///
                 E`check_row' = `lower_bound' F`check_row' = `upper_bound' ///
                 G`check_row' = `valid_n' H`check_row' = "`check_status'"
        local ++check_row
    }
}


/*******************************************************************************
*  SECTION 4. TABLE SPECIFICATIONS
*  For each table: sheet title, components (pooled = four components
*  together), breakdowns (overall = no breakdown) and indicators.
*******************************************************************************/

local table_list "T1_Satisfaction T2_Complaint_pathway T3_Reporting_safety T4_CLAS_county T5_Round T6_Sex T7_Inputs_decisions"

local title_T1_Satisfaction       "Overall satisfaction (G07), satisfied or very satisfied, by component and sex (PAP secondary)"
local components_T1_Satisfaction  "ssb lipw clas sct"
local breakdowns_T1_Satisfaction  "overall sex_bin"
local indicators_T1_Satisfaction  "satisfied_overall"

local title_T2_Complaint_pathway      "Complaint pathway: knowledge, use, access when needed, resolution (PAP secondary GRM family; pooled row exploratory)"
local components_T2_Complaint_pathway "ssb lipw clas sct pooled"
local breakdowns_T2_Complaint_pathway "overall"
local indicators_T2_Complaint_pathway "knows_complaint_channel used_complaint_channel channel_access_easy channel_not_accessible complaint_resolved_any"

local title_T3_Reporting_safety      "Feeling safe reporting fraud, misconduct, harassment or abuse (G06), by component and sex (PAP secondary)"
local components_T3_Reporting_safety "ssb lipw clas sct"
local breakdowns_T3_Reporting_safety "overall sex_bin"
local indicators_T3_Reporting_safety "safe_to_report_misconduct unsafe_to_report_misconduct"

local title_T4_CLAS_county      "CLAS delivery and reported change by county (exploratory; county overlaps interviewer assignment)"
local components_T4_CLAS_county "clas"
local breakdowns_T4_CLAS_county "overall county_code"
local indicators_T4_CLAS_county "p1 p2 no_schedule_communicated worse_off_basic_needs clas_full_input_package"

local title_T5_Round      "Receipt, timeliness, basic-needs change and complaint knowledge by represented round (PAP secondary subgroup, descriptive)"
local components_T5_Round "ssb lipw clas"
local breakdowns_T5_Round "round_num"
local indicators_T5_Round "p1 p2 p4 knows_complaint_channel"

local title_T6_Sex      "Primary outcomes P1 to P4 by respondent sex (PAP primary subgroup)"
local components_T6_Sex "ssb lipw clas sct"
local breakdowns_T6_Sex "sex_bin"
local indicators_T6_Sex "p1 p2 p3 p4"

local title_T7_Inputs_decisions      "Own or joint decision on use of support (R06B) by sex; CLAS input package, harvest and expected continuation (descriptive)"
local components_T7_Inputs_decisions "ssb lipw clas sct"
local breakdowns_T7_Inputs_decisions "overall sex_bin"
local indicators_T7_Inputs_decisions "own_or_joint_decision clas_full_input_package clas_harvest_completed clas_continuation_likely"

local score_indicators "p3 p4"


/*******************************************************************************
*  SECTION 5. ESTIMATION AND EXPORT
*  One row per component, breakdown group and indicator. Display text applies
*  the small-cell rules; numeric cells are left empty when suppressed.
*******************************************************************************/

foreach table_name of local table_list {

    putexcel set "$REALISE_OUTPUT_WORKBOOK", sheet("`table_name'") modify
    putexcel A1 = "`title_`table_name''"
    putexcel A2 = "Component" B2 = "Breakdown" C2 = "Group" D2 = "Indicator" ///
             E2 = "Weighted estimate" F2 = "95% CI lower" G2 = "95% CI upper" ///
             H2 = "Valid n" I2 = "Display text" J2 = "Complementary cell below 10"
    local table_row = 3

    foreach component of local components_`table_name' {

        use "`stacked_analysis_file'", clear
        if "`component'" != "pooled" keep if component_name == "`component'"
        svyset design_psu [pweight = final_weight], strata(design_stratum) fpc(N_h) singleunit(centered)

        foreach breakdown of local breakdowns_`table_name' {

            if "`breakdown'" == "overall" local group_values "1"
            else quietly levelsof `breakdown', local(group_values)

            foreach group_value of local group_values {

                if "`breakdown'" == "overall" {
                    local group_condition "1"
                    local group_label "All"
                }
                else {
                    local group_condition "`breakdown' == `group_value'"
                    local group_label : label (`breakdown') `group_value'
                }

                foreach indicator of local indicators_`table_name' {

                    **** Domain flag and zero-filled copy keep every interview in the design
                    capture drop domain_flag analysis_value
                    generate byte domain_flag = (`group_condition') & !missing(`indicator')
                    generate double analysis_value = cond(missing(`indicator'), 0, `indicator')
                    quietly count if domain_flag
                    local valid_n = r(N)
                    if `valid_n' == 0 continue

                    **** Complementary-cell flag for 0/1 indicators: either answer given by 1 to 9 respondents
                    quietly count if domain_flag & analysis_value == 1
                    local positive_count = r(N)
                    local negative_count = `valid_n' - `positive_count'
                    local complementary_flag "No"
                    if inrange(`positive_count', 1, 9) | inrange(`negative_count', 1, 9) local complementary_flag "Yes"

                    quietly svy, subpop(domain_flag): mean analysis_value
                    matrix indicator_estimate = r(table)
                    local point_estimate = indicator_estimate[1, 1]
                    local standard_error = indicator_estimate[2, 1]
                    local t_critical     = invttail(e(df_r), 0.025)
                    local is_score : list indicator in score_indicators

                    if `is_score' {
                        local complementary_flag "Not applicable"
                        local reported_value = `point_estimate'
                        local lower_bound = `point_estimate' - `t_critical' * `standard_error'
                        local upper_bound = `point_estimate' + `t_critical' * `standard_error'
                        local display_text = string(`reported_value', "%5.2f") + " [" + ///
                            string(`lower_bound', "%5.2f") + ", " + string(`upper_bound', "%5.2f") + "]"
                    }
                    else {
                        local reported_value = 100 * `point_estimate'
                        if `point_estimate' > 0 & `point_estimate' < 1 & `standard_error' > 0 {
                            local logit_se    = `standard_error' / (`point_estimate' * (1 - `point_estimate'))
                            local lower_bound = 100 * invlogit(logit(`point_estimate') - `t_critical' * `logit_se')
                            local upper_bound = 100 * invlogit(logit(`point_estimate') + `t_critical' * `logit_se')
                            local display_text = string(`reported_value', "%5.1f") + "% [" + ///
                                string(`lower_bound', "%5.1f") + ", " + string(`upper_bound', "%5.1f") + "]"
                        }
                        else {
                            local lower_bound = .
                            local upper_bound = .
                            local display_text = string(`reported_value', "%5.1f") + "% (CI not estimable)"
                        }
                    }

                    **** Small-cell rules
                    if `valid_n' < `minimum_reported_n' {
                        local display_text "Suppressed"
                        local reported_value = .
                        local lower_bound = .
                        local upper_bound = .
                    }
                    else if `valid_n' < `minimum_firm_n' {
                        local display_text "`display_text' (indicative)"
                    }
                    local display_text "`display_text' (n=`valid_n')"

                    putexcel A`table_row' = "`component'" B`table_row' = "`breakdown'" ///
                             C`table_row' = "`group_label'" D`table_row' = "`indicator'" ///
                             H`table_row' = `valid_n' I`table_row' = "`display_text'" ///
                             J`table_row' = "`complementary_flag'"
                    if `reported_value' < . putexcel E`table_row' = `reported_value'
                    if `lower_bound' < .    putexcel F`table_row' = `lower_bound'
                    if `upper_bound' < .    putexcel G`table_row' = `upper_bound'
                    local ++table_row
                }
            }
        }
    }
}

**** End of file
