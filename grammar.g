%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.WS
%import common.CNAME -> IDENTIFIER

%ignore WS

start: (structs | COMMENT)*
structs: structs struct | struct
struct: "Function" IDENTIFIER i_fns "End" "Function"

COMMENT: "'" /.*[^\\n]/

i_st_opn_db: "Set" IDENTIFIER "=" IDENTIFIER "." "OpenDatabase" "(" STRING_LITERAL ")"
i_st_nts: i_st_nts i_st_nt | i_st_nt
i_st_nt: "Set" IDENTIFIER "=" "Nothing"

HAVE_RECORDS: "haveRecords" | "HaveRecords"
i_cnd_dcl: "If" HAVE_RECORDS "(" STRING_LITERAL ")" ("And" HAVE_RECORDS "(" STRING_LITERAL ")")* "Then"
i_cnd_tnl: "End" ("If" | "IF" | "if")

i_fns: i_fns i_fn | i_fn
i_fn: i_cnd_dcl i_st_opn_db i_st_fn i_st_nts i_cnd_tnl

i_st_mtd: IDENTIFIER "." s_st_fns
i_st_fn: "Set" IDENTIFIER "=" i_st_mtd

s_st_fns: d_summ | d_join | d_export | d_extract | d_connect | d_sort | d_dup_excl | d_tbl_mgmt
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


d_extract: "Extraction" (IDENTIFIER "." s_extract_opts | i_db_name)+
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

d_tbl_mgmt: "TableManagement" e_tbl_mgmt_st s_tbl_mgmt_opts IDENTIFIER "." (d_add_col | d_rename_col) (IDENTIFIER "." i_perf_tsk)?
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
