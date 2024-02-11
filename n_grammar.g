%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.WS
%import common.CNAME -> IDENTIFIER

%ignore WS

COMMENT: "'" /.*[^\\n]/
FUNCTION: "Function" | "FUNCTION" | "Function"
END: "End" | "END" | "end"
IF: "If" | "IF" | "if"
ELSE: "Else" | "ELSE" | "else"
HAVERECORDS: "haveRecords" | "HaveRecords" | "HAVERECORDS" | "haverecords"
AND: "And" | "AND" | "and"
THEN: "Then" | "THEN" | "then"

DOUBLE_QUOTE : "\""

s_bools: TRUE | FALSE
TRUE: "TRUE" | "true"
FALSE: "FALSE" | "false"

prog: start
start: iss_opts
iss_opts: iss_opts iss_opt | iss_opt | 
iss_opt: struct | COMMENT

struct: struct_cond_decl struct_opts struct_cond_tnl
struct_opts: struct_opts struct_opt | struct_opt
struct_opt: i_fns | st_nts
struct_cond_decl: FUNCTION IDENTIFIER 
struct_cond_tnl: END FUNCTION

st_nts: st_nts st_nt | st_nt
st_nt: "Set" st_opts "=" "Nothing"
st_opts: "field" | "task" | "db"

i_fns: bp_std_fns | std_fns

bp_std_fns: have_records_check_decl std_fns_decl have_records_check_tnl have_records_check_else?
have_records_check_decl: IF have_records_opts THEN
have_records_opts: have_records_opts have_records_opt | have_records_opt
have_records_opt: have_records | AND
have_records: HAVERECORDS "(" STRING_LITERAL ")"
have_records_check_tnl: END IF
have_records_check_else: ELSE bp_method_opts
bp_method_opts: "NORESULTSLOG" "(" STRING_LITERAL ")"

std_fns_decl: st_open_db st_fn
st_open_db: "Set" "db" "=" "Client" "." "OpenDatabase" "(" STRING_LITERAL ")"
st_fn: "Set" "task" "=" "db" "." std_fns_opts

std_fns: std_fns_decl

std_fns_opts: (d_summarize) st_nts?

d_summarize: "Summarization" e_summarize_opts
e_summarize_opts: e_summarize_opts e_summarize_opt | e_summarize_opt
e_summarize_opt: e_summarize_task_opts | e_summ_db_name
e_summarize_task_opts: "task" "." s_summarize_task_opts
s_summarize_task_opts: e_add_fields_to_summarize | e_add_fields_to_total | e_summ_criteria | e_summ_output_db_name | e_summ_create_percent_field | e_summ_statistics_to_include | e_perform_task 
e_add_fields_to_summarize: e_add_fields_to_summarize e_add_field_to_summarize | e_add_field_to_summarize
e_add_field_to_summarize: "AddFieldToSummarize" STRING_LITERAL
e_add_fields_to_total: e_add_fields_to_total e_add_field_to_total | e_add_field_to_total
e_add_field_to_total: "AddFieldToTotal" STRING_LITERAL
e_summ_criteria: "Criteria" "=" STRING_LITERAL
e_summ_output_db_name: "OutputDBName" "=" (STRING_LITERAL | IDENTIFIER)
e_summ_create_percent_field: "CreatePercentField" "=" s_bools
e_summ_statistics_to_include: "StatisticsToInclude" "=" s_stats_opts
e_perform_task: "PerformTask"
e_summ_db_name: "dbName" "=" STRING_LITERAL
s_stats_opts: SM_SUM
SM_SUM: "SM_SUM"



