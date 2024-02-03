%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.WS
%import common.CNAME -> IDENTIFIER

%ignore WS

COMMENT: "'" /.*[^\\n]/

prog: start
start: iss_opts
iss_opts: iss_opts iss_opt | iss_opt | 
iss_opt: structs | COMMENT | vars_t | subs

structs: structs struct | struct
struct: struct_cond_decl struct_opts "End" "Function"
struct_opts: i_fns | i_cleanup | i_log_file
struct_cond_decl: "Function" IDENTIFIER 

vars_t: vars_t var | var
var: var_type_decl_opts IDENTIFIER var_type_opts
var_type_decl_opts: const_v | global_v | reg_v
var_type_opts: asgn_opts | decl_opts
decl_opts: "As" type_opts
asgn_opts: "=" val_opts
const_v: "Const"
global_v: "Global"
reg_v: "Dim"
val_opts: STRING_LITERAL | INT
type_opts: type_str | type_obj | type_double
type_str: "String" 
type_obj: "Object" 
type_double: "Double"

subs: sub_cond_decl subroutine_opts "End" "Sub"
sub_cond_decl: "Sub" IDENTIFIER
subroutine_opts: subroutine_opts subroutine_opt | subroutine_opt
subroutine_opt: sub_ignore_warn | sub_set | sub_on_err  | COMMENT | sub_client_opts | IDENTIFIER | final_routine | call_fns_opts
sub_ignore_warn: "Ignorewarning" "(" s_bools ")"
sub_set: "Set" IDENTIFIER "=" "CreateObject" "(" STRING_LITERAL ")"
sub_on_err: "On" "Error" "GoTo" IDENTIFIER
sub_client_opts: "Client" "." (sub_close_all | sub_quit)
sub_close_all: "CloseAll"
sub_quit: "Quit"
final_routine: IDENTIFIER ":" cond_err_decl
call_fns_opts: "Call" IDENTIFIER call_param?
call_param: "(" call_param_opts ")"
call_param_opts: call_param_opts call_param_opt | call_param_opt
call_param_opt: STRING_LITERAL | "," | "&" | IDENTIFIER

cond_err_decl: "If" cond_err_decl_opts "Then" cond_err_opts "Else" cond_err_opts "End" "If"
cond_err_decl_opts: cond_err_decl_opts cond_err_decl_opt | cond_err_decl_opt
cond_err_decl_opt: err_method | "<>" | "Or" | STRING_LITERAL | INT
err_method: "err" "." err_method_opts
err_method_opts: "description" | "number"

cond_err_opts: cond_err_opts cond_err_opt | cond_err_opt
cond_err_opt: IDENTIFIER "=" cond_err_asgn_opts
cond_err_asgn_opts: cond_err_asgn_opts cond_err_asgn_opt | cond_err_asgn_opt
cond_err_asgn_opt: err_method | STRING_LITERAL | INT | "&"

i_st_opn_db: "Set" IDENTIFIER "=" "Client" "." "OpenDatabase" "(" STRING_LITERAL ")"
i_st_nts: i_st_nts i_st_nt | i_st_nt
i_st_nt: "Set" IDENTIFIER "=" "Nothing"

i_cnd_dcl: "If" i_cnd_dcl_opts "Then"
i_cnd_dcl_opts: i_cnd_dcl_opts i_cnd_dcl_opt | i_cnd_dcl_opt
i_cnd_dcl_opt: i_cnd_hav | "And"
i_cnd_hav: "haveRecords" "(" STRING_LITERAL ")"
i_cnd_els: "Else" call_fns_opts
i_cnd_tnl: "End" "If"

i_fns: i_fns i_fn | i_fn
i_fn: i_cnd_dcl? i_st_opn_db i_st_fn i_st_nts? i_cnd_els? i_cnd_tnl?

i_st_mtd: IDENTIFIER "." s_st_fns
i_st_fn: "Set" IDENTIFIER "=" i_st_mtd

s_st_fns: d_summ | d_join | d_export | d_extract | d_connect | d_sort | d_dup_excl | d_tbl_mgmt | COMMENT
s_bools: "FALSE" | "TRUE" | "True" | "False"

d_summ: "Summarization" (IDENTIFIER "." s_summ_opts | i_db_name)+
s_summ_opts: "AddFieldToSummarize" STRING_LITERAL | i_add_flds_to_inc | e_out_db_name | e_crt_per_fld | e_use_fld_occurr | i_perf_tsk

e_out_db_name: "OutputDBName" "=" IDENTIFIER
e_crt_per_fld: "CreatePercentField" "=" s_bools
e_use_fld_occurr: "UseFieldFromFirstOccurrence" "=" s_bools

i_perf_tsk: "PerformTask"
i_db_name: IDENTIFIER "=" STRING_LITERAL
i_add_flds_to_inc: i_add_flds_to_inc i_add_fld_to_inc | i_add_fld_to_inc
i_add_fld_to_inc: "AddFieldToInc" STRING_LITERAL
i_crt_virt_db: "CreateVirtualDatabase" "=" s_bools
i_add_keys: i_add_keys i_add_key | i_add_key
i_add_key: "AddKey" STRING_LITERAL "," STRING_LITERAL

d_join: "JoinDatabase" (IDENTIFIER "." s_join_opts | i_db_name)+
s_join_opts: "FileToJoin" STRING_LITERAL | e_incl_all_p_fld | e_add_s_fld | e_add_p_fld | e_incl_all_s_fld | e_add_mtch | i_crt_virt_db | e_perf_task_join
e_incl_all_p_fld: "IncludeAllPFields"
e_incl_all_s_fld: "IncludeAllSFields"
e_add_s_fld: "AddSFieldToInc" STRING_LITERAL
e_add_p_fld: "AddPFieldToInc" STRING_LITERAL
e_add_mtch: "AddMatchKey" STRING_LITERAL "," STRING_LITERAL "," STRING_LITERAL
e_perf_task_join: "PerformTask" IDENTIFIER "," STRING_LITERAL "," s_tsk_opts
s_tsk_opts: (WI_JOIN_MATCH_ONLY | WI_JOIN_ALL_IN_PRIM | WI_JOIN_NOC_SEC_MATCH)
WI_JOIN_MATCH_ONLY: "WI_JOIN_MATCH_ONLY"
WI_JOIN_ALL_IN_PRIM: "WI_JOIN_ALL_IN_PRIM"
WI_JOIN_NOC_SEC_MATCH: "WI_JOIN_NOC_SEC_MATCH"


d_extract: "Extraction" d_extract_opts
d_extract_opts: d_extract_opts d_extract_opt | d_extract_opt
d_extract_opt: IDENTIFIER "." s_extract_opts | i_db_name
s_extract_opts: i_add_flds_to_inc | e_add_extraction | i_crt_virt_db | e_perf_task_extract
e_add_extraction: "AddExtraction" IDENTIFIER "," STRING_LITERAL "," STRING_LITERAL
e_perf_task_extract: "PerformTask" INT "," IDENTIFIER "." IDENTIFIER

d_connect: "VisualConnector" e_add_databases (IDENTIFIER "." s_connect_opts | i_db_name)+
s_connect_opts: e_mast_db | e_apnd_db_names | e_incl_all_p_recs | e_add_rel | i_crt_virt_db | i_perf_tsk | e_add_flds_to_incl | e_out_database_name
e_mast_db: "MasterDatabase" "=" IDENTIFIER
e_apnd_db_names: "AppendDatabaseNames" "=" s_bools
e_incl_all_p_recs: "IncludeAllPrimaryRecords" "=" s_bools
e_out_database_name: "OutputDatabaseName" "=" IDENTIFIER
e_add_rel: "AddRelation" IDENTIFIER "," STRING_LITERAL "," IDENTIFIER "," STRING_LITERAL
e_add_flds_to_incl: e_add_flds_to_incl e_add_fld_to_incl | e_add_fld_to_incl
e_add_fld_to_incl: "AddFieldToInclude" IDENTIFIER "," STRING_LITERAL
e_add_databases: e_add_databases e_add_database | e_add_database
e_add_database: IDENTIFIER "=" IDENTIFIER "." "AddDatabase" "(" STRING_LITERAL ")"

d_sort: "Sort" (IDENTIFIER "." s_sort_opts | i_db_name)+
s_sort_opts: i_add_keys | e_perf_task_sort
e_perf_task_sort: "PerformTask" IDENTIFIER

d_export: "ExportDatabase" (IDENTIFIER "." s_export_opts | i_db_name)+
s_export_opts: i_add_flds_to_inc | e_perf_task_export
e_perf_task_export: "PerformTask" IDENTIFIER "." IDENTIFIER "&" STRING_LITERAL "," STRING_LITERAL ","  STRING_LITERAL "," INT "," IDENTIFIER "." IDENTIFIER "," IDENTIFIER

d_dup_excl: "DupKeyExclusion" (IDENTIFIER "." s_dup_excl_opts | i_db_name)+
s_dup_excl_opts: e_incl_all_flds | i_add_keys | e_diff_fld | i_crt_virt_db | e_perf_task_excl
e_incl_all_flds: "IncludeAllFields"
e_diff_fld: "DifferentField" "=" STRING_LITERAL
e_perf_task_excl: "PerformTask" IDENTIFIER "," STRING_LITERAL

d_tbl_mgmt: "TableManagement" (e_tbl_mgmt_st s_tbl_mgmt_opts IDENTIFIER "." (d_add_col | d_rename_col) (IDENTIFIER "." i_perf_tsk)?)+
e_tbl_mgmt_st: "Set" IDENTIFIER "=" IDENTIFIER "." "TableDef" "." "NewField"
s_tbl_mgmt_opts: s_tbl_mgmt_opts s_tbl_mgmt_opt | s_tbl_mgmt_opt
s_tbl_mgmt_opt: e_tbl_mgmt_name | e_tbl_mgmt_desc | e_tbl_mgmt_type | e_tbl_mgmt_eqn | e_tbl_mgmt_len
e_tbl_mgmt_name: IDENTIFIER "." "Name" "=" STRING_LITERAL
e_tbl_mgmt_desc: IDENTIFIER "." "Description" "=" STRING_LITERAL
e_tbl_mgmt_eqn: IDENTIFIER "." "Equation" "=" STRING_LITERAL
e_tbl_mgmt_len: IDENTIFIER "." "Length" "=" INT
e_tbl_mgmt_type: IDENTIFIER "." "Type" "=" s_tbl_mgmt_types
s_tbl_mgmt_types: WI_CHAR_FIELD | WI_TIME_FIELD | WI_DATE_FIELD | WI_VIRT_CHAR
WI_CHAR_FIELD: "WI_CHAR_FIELD"
WI_TIME_FIELD: "WI_TIME_FIELD"
WI_DATE_FIELD: "WI_DATE_FIELD"
WI_VIRT_CHAR: "WI_VIRT_CHAR"

d_add_col: "AppendField" IDENTIFIER
d_rename_col: "ReplaceField" STRING_LITERAL "," IDENTIFIER

i_cleanup: i_cleanup_opts
i_cleanup_opts: i_cleanup_opts i_cleanup_opt | i_cleanup_opt
i_cleanup_opt: "DeleteFile" "(" STRING_LITERAL ")" | COMMENT

i_log_file: i_log_file_decl sub_on_err i_log_file_cond_err i_log_file_vars i_log_file_cond i_log_file_set i_log_file_set_cond
i_log_file_decl: "(" i_by_vals ")"
i_by_vals: i_by_vals i_by_val | i_by_val
i_by_val: "ByVal" IDENTIFIER "As" type_opts | ","
i_log_file_cond_err: "If" "e_debug" "<>" "True" "Then" "Exit" "Sub"
i_log_file_vars: i_log_file_vars i_log_file_var | i_log_file_var
i_log_file_var: "Dim" IDENTIFIER "As" type_opts

i_log_file_cond: "If" "(" "Len" "(" "e_logfilename" ")" ">" INT ")" "Then" IDENTIFIER "=" "e_logfilename" "&" STRING_LITERAL "Else" IDENTIFIER "=" STRING_LITERAL
i_log_file_set: "Set" IDENTIFIER "=" "Client" "." "ProjectManagement"
i_log_file_set_cond: "If" "Not" IDENTIFIER "." "DoesDatabaseExist" "(" IDENTIFIER ")" "Then" i_log_file_set_cond_opts "End" "If"
i_log_file_set_cond_opts: s_log_file_set_opts
s_log_file_set_opts: s_log_file_set_opts s_log_file_set_opt | s_log_file_set_opt
s_log_file_set_opt: "Set" IDENTIFIER "=" s_set_opts
s_set_opts: "Nothing" | i_log_client | i_log_new_fld
i_log_client: "Client" "." s_log_client_opts
s_log_client_opts: i_log_new_db | i_log_new_tbl
i_log_new_fld: "NewTable" "." "NewField"
i_log_new_db: "NewDatabase" "(" s_log_new_db_opts ")" 
s_log_new_db_opts: s_log_new_db_opts s_log_new_db_opt | s_log_new_db_opt
s_log_new_db_opt: IDENTIFIER | "," | STRING_LITERAL
i_log_new_tbl: "NewTableDef"