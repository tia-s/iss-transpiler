%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.FLOAT -> FLOAT
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
TRUE: "TRUE" | "true" | "True"
FALSE: "FALSE" | "false" | "False"

prog: start
start: iss_opts
iss_opts: iss_opts iss_opt | iss_opt | 
iss_opt: struct | COMMENT

iss_methods: "@" iss_methods_opts
iss_methods_opts: iss_match_method

iss_match_method: "Match" "(" iss_match_method_opts ")"
iss_match_method_opts: iss_match_method_opts iss_match_method_opt | iss_match_method_opt
iss_match_method_opt: IDENTIFIER | STRING_LITERAL | "," 

iss_not_equal: IDENTIFIER "<>" iss_equate_opts
iss_equals: IDENTIFIER "==" iss_equate_opts
iss_equate_opts: DOUBLE_QUOTE STRING_LITERAL DOUBLE_QUOTE | INT | FLOAT

struct: struct_cond_decl struct_opts struct_cond_tnl
struct_opts: struct_opts struct_opt | struct_opt
struct_opt: i_fns | st_nts
struct_cond_decl: FUNCTION IDENTIFIER 
struct_cond_tnl: END FUNCTION

st_nts: st_nts st_nt | st_nt
st_nt: "Set" st_opts "=" "Nothing"
st_opts: "field" | "task" | "db"

i_fns: bp_std_fns | std_fns | d_cleanup

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

std_fns_opts: (d_summarize | d_join | d_extract | d_export | d_tbl_manage | d_visual_connect) st_nts?

d_summarize: "Summarization" e_summarize_opts
e_summarize_opts: e_summarize_opts e_summarize_opt | e_summarize_opt
e_summarize_opt: e_summarize_task_opts | e_summ_db_name
e_summarize_task_opts: "task" "." s_summarize_task_opts
s_summarize_task_opts: e_add_fields_to_summarize | e_add_fields_to_total | e_add_fields_to_inc | e_summ_criteria | e_summ_output_db_name | e_summ_create_percent_field | e_summ_statistics_to_include | e_perform_task 
e_add_fields_to_summarize: e_add_fields_to_summarize e_add_field_to_summarize | e_add_field_to_summarize
e_add_field_to_summarize: "AddFieldToSummarize" STRING_LITERAL
e_add_fields_to_total: e_add_fields_to_total e_add_field_to_total | e_add_field_to_total
e_add_field_to_total: "AddFieldToTotal" STRING_LITERAL
e_add_fields_to_inc: e_add_fields_to_inc e_add_field_to_inc | e_add_field_to_inc
e_add_field_to_inc: "AddFieldToInc" STRING_LITERAL
e_summ_criteria: "Criteria" "=" DOUBLE_QUOTE s_summ_criteria_opts DOUBLE_QUOTE
s_summ_criteria_opts: s_summ_criteria_opts s_summ_criteria_opt | s_summ_criteria_opt
s_summ_criteria_opt: iss_methods | ".AND." | iss_not_equal | iss_equals
e_summ_output_db_name: "OutputDBName" "=" (STRING_LITERAL | IDENTIFIER)
e_summ_create_percent_field: "CreatePercentField" "=" s_bools
e_summ_statistics_to_include: "StatisticsToInclude" "=" s_stats_opts
e_perform_task: "PerformTask"
e_summ_db_name: "dbName" "=" STRING_LITERAL
s_stats_opts: s_stats_opts s_stats_opt | s_stats_opt
s_stats_opt: SM_SUM | SM_AVERAGE | "+"
SM_SUM: "SM_SUM"
SM_AVERAGE: "SM_AVERAGE"

d_join: "JoinDatabase" e_join_opts
e_join_opts: e_join_opts e_join_opt | e_join_opt
e_join_opt: e_join_task_opts | e_join_db_name
e_join_task_opts: "task" "." s_join_task_opts
s_join_task_opts: e_file_to_join | e_include_all_p_fields | e_include_all_s_fields | e_add_s_fields_to_inc | e_add_p_fields_to_inc | e_add_match_key | e_join_create_virt_database | e_join_perform_task
e_file_to_join: "FileToJoin" STRING_LITERAL
e_include_all_p_fields: "IncludeAllPFields"
e_include_all_s_fields: "IncludeAllSFields"
e_add_s_fields_to_inc: e_add_s_fields_to_inc e_add_s_field_to_inc | e_add_s_field_to_inc
e_add_s_field_to_inc: "AddSFieldToInc" STRING_LITERAL
e_add_p_fields_to_inc: e_add_p_fields_to_inc e_add_p_field_to_inc | e_add_p_field_to_inc
e_add_p_field_to_inc: "AddPFieldToInc" STRING_LITERAL
e_add_match_key: "AddMatchKey" STRING_LITERAL "," STRING_LITERAL "," STRING_LITERAL
e_join_create_virt_database: "CreateVirtualDatabase" "=" s_bools
e_join_perform_task: "PerformTask" IDENTIFIER "," STRING_LITERAL "," s_join_types
s_join_types: WI_JOIN_ALL_IN_PRIM | WI_JOIN_MATCH_ONLY | WI_JOIN_NOC_SEC_MATCH
WI_JOIN_ALL_IN_PRIM: "WI_JOIN_ALL_IN_PRIM"
WI_JOIN_MATCH_ONLY: "WI_JOIN_MATCH_ONLY"
WI_JOIN_NOC_SEC_MATCH: "WI_JOIN_NOC_SEC_MATCH"
e_join_db_name: "dbName" "=" STRING_LITERAL

d_extract: "Extraction" e_extract_opts
e_extract_opts: e_extract_opts e_extract_opt | e_extract_opt
e_extract_opt: e_extract_task_opts | e_extract_db_name
e_extract_task_opts: "task" "." s_extract_task_opts
s_extract_task_opts: e_extract_include_all_fields | e_extract_add_fields_to_inc | e_add_extraction | e_extract_create_virt_database | e_extract_perform_task
e_extract_include_all_fields: "IncludeAllFields"
e_extract_add_fields_to_inc: e_extract_add_fields_to_inc e_extract_add_field_to_inc | e_extract_add_field_to_inc
e_extract_add_field_to_inc: "AddFieldToInc" STRING_LITERAL
e_add_extraction: "AddExtraction" "dbName" "," STRING_LITERAL "," s_extract_filter_opts
s_extract_filter_opts: s_extract_filter_opts s_extract_filter_opt | s_extract_filter_opt
s_extract_filter_opt: STRING_LITERAL | "&" | IDENTIFIER
e_extract_create_virt_database: "CreateVirtualDatabase" "=" s_bools
e_extract_perform_task: "PerformTask" INT "," "db" "." "Count"
e_extract_db_name: "dbName" "=" STRING_LITERAL

d_export: "ExportDatabase" e_export_opts
e_export_opts: e_export_opts e_export_opt | e_export_opt
e_export_opt: e_export_task_opts | e_export_eqn
e_export_task_opts: "task" "." s_export_task_opts
s_export_task_opts: e_export_add_fields_to_inc | e_export_perform_task
e_export_add_fields_to_inc: e_export_add_fields_to_inc e_export_add_field_to_inc | e_export_add_field_to_inc
e_export_add_field_to_inc: "AddFieldToInc" STRING_LITERAL
e_export_perform_task: "PerformTask" "Client" "." "WorkingDirectory" s_export_perform_task_opts
s_export_perform_task_opts: s_export_perform_task_opts s_export_perform_task_opt | s_export_perform_task_opt
s_export_perform_task_opt: "&" | "," | INT | STRING_LITERAL | IDENTIFIER "." IDENTIFIER | "eqn" | IDENTIFIER
e_export_eqn: "eqn" "=" STRING_LITERAL

d_cleanup: e_cleanup_task_opts
e_cleanup_task_opts: e_cleanup_task_opts e_cleanup_task_opt | e_cleanup_task_opt
e_cleanup_task_opt: e_cleanup_delete_files
e_cleanup_delete_files: e_cleanup_delete_files e_cleanup_delete_file | e_cleanup_delete_file
e_cleanup_delete_file: "DeleteFile" "(" STRING_LITERAL ")"

d_tbl_manage: d_tbl_manage d_tbl_manage_ | d_tbl_manage_
d_tbl_manage_: ("Set" "task" "=" "db" ".")? "TableManagement" e_tbl_mgmt_st_field e_tbl_mgmt_opts
e_tbl_mgmt_st_field: "Set" "field" "=" "db" "." "TableDef" "." "NewField"
e_tbl_mgmt_opts: e_tbl_mgmt_opts e_tbl_mgmt_opt | e_tbl_mgmt_opt
e_tbl_mgmt_opt: e_tbl_mgmt_field_opts | e_tbl_mgmt_task_opts
e_tbl_mgmt_field_opts: "field" "." s_tbl_mgmt_field_opts
s_tbl_mgmt_field_opts: e_tbl_mgmt_name | e_tbl_mgmt_desc | e_tbl_mgmt_len | e_tbl_mgmt_type | e_tbl_mgmt_eqn | e_tbl_mgmt_decimals
e_tbl_mgmt_name: "Name" "=" STRING_LITERAL
e_tbl_mgmt_desc: "Description" "=" STRING_LITERAL
e_tbl_mgmt_len: "Length" "=" INT
e_tbl_mgmt_type: "Type" "=" s_tbl_mgmt_types
e_tbl_mgmt_decimals: "Decimals" "=" INT
e_tbl_mgmt_eqn: "Equation" "=" STRING_LITERAL
s_tbl_mgmt_types: WI_CHAR_FIELD | WI_TIME_FIELD | WI_DATE_FIELD | WI_VIRT_CHAR | WI_VIRT_DATE | WI_NUM_FIELD
WI_CHAR_FIELD: "WI_CHAR_FIELD"
WI_TIME_FIELD: "WI_TIME_FIELD"
WI_DATE_FIELD: "WI_DATE_FIELD"
WI_VIRT_CHAR: "WI_VIRT_CHAR"
WI_VIRT_DATE: "WI_VIRT_DATE"
WI_NUM_FIELD: "WI_NUM_FIELD"
e_tbl_mgmt_task_opts: "task" "." s_tbl_mgmt_task_opts
s_tbl_mgmt_task_opts: e_tbl_mgmt_append_field | e_tbl_mgmt_replace_field | e_tbl_mgmt_perform_task
e_tbl_mgmt_append_field: "AppendField" "field"
e_tbl_mgmt_replace_field: "ReplaceField" STRING_LITERAL "," "field"
e_tbl_mgmt_perform_task: "PerformTask"

d_visual_connect: "VisualConnector" e_visual_connect_opts
e_visual_connect_opts: e_visual_connect_opts e_visual_connect_opt | e_visual_connect_opt
e_visual_connect_opt: e_visual_connect_task_opts | e_visual_connect_db_name | e_visual_connect_add_assgns
e_visual_connect_task_opts: "task" "." s_visual_connect_task_opts
s_visual_connect_task_opts: e_visual_connect_add_fields_to_include | e_visual_connect_add_relations | e_visual_connect_master_database | e_visual_connect_append_database_names | e_visual_conenct_include_all_primary_recs | e_visual_connect_add_database | e_visual_connect_create_virt_database | e_visual_connect_output_db_name | e_visual_connect_perf_task
e_visual_connect_add_fields_to_include: e_visual_connect_add_fields_to_include e_visual_connect_add_field_to_include | e_visual_connect_add_field_to_include
e_visual_connect_add_field_to_include: "AddFieldToInclude" IDENTIFIER "," STRING_LITERAL
e_visual_connect_add_relations: e_visual_connect_add_relations e_visual_connect_add_relation | e_visual_connect_add_relation
e_visual_connect_add_relation: "AddRelation" s_visual_connect_add_relation_opts
s_visual_connect_add_relation_opts: s_visual_connect_add_relation_opts s_visual_connect_add_relation_opt | s_visual_connect_add_relation_opt
s_visual_connect_add_relation_opt: IDENTIFIER | STRING_LITERAL | ","  | e_visual_connect_task_opts
e_visual_connect_master_database: "MasterDatabase" "=" (IDENTIFIER | e_visual_connect_task_opts)
e_visual_connect_append_database_names: "AppendDatabaseNames" "=" s_bools
e_visual_conenct_include_all_primary_recs: "IncludeAllPrimaryRecords" "=" s_bools
e_visual_connect_add_database: "AddDatabase" "(" STRING_LITERAL ")"
e_visual_connect_create_virt_database: "CreateVirtualDatabase" "=" s_bools
e_visual_connect_output_db_name: "OutputDatabaseName" "=" (IDENTIFIER | STRING_LITERAL)
e_visual_connect_perf_task: "PerformTask"
e_visual_connect_db_name: "dbName" "=" STRING_LITERAL
e_visual_connect_add_assgns: e_visual_connect_add_assgns e_visual_connect_add_assgn | e_visual_connect_add_assgn
e_visual_connect_add_assgn: IDENTIFIER "=" e_visual_connect_task_opts