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


prog: start
start: iss_opts
iss_opts: iss_opts iss_opt | iss_opt | 
iss_opt: struct | COMMENT

struct: struct_cond_decl struct_opts struct_cond_tnl
struct_opts: struct_opts struct_opt | struct_opt
struct_opt: i_fns
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

std_fns_opts: d_tbl_mgmts
d_tbl_mgmts: d_tbl_mgmts d_tbl_mgmt | d_tbl_mgmt
d_tbl_mgmt: "TableManagement" e_tbl_mgmt_st e_tbl_opts | e_tbl_mgmt_inner
e_tbl_mgmt_st: "Set" "field" "=" "db" "." "TableDef" "." "NewField"
e_tbl_opts: e_tbl_opts e_tbl_opt | e_tbl_opt
e_tbl_opt: e_tbl_mgmt_opts | e_tbl_mgmt_task_opts | e_tbl_mgmt_set_tsk_opts | e_tbl_mgmt_set_oth_nts
e_tbl_mgmt_opts: e_tbl_mgmt_opts e_tbl_mgmt_opt | e_tbl_mgmt_opt
e_tbl_mgmt_opt: "field" "." (e_tbl_mgmt_name | e_tbl_mgmt_desc | e_tbl_mgmt_type | e_tbl_mgmt_eqn | e_tbl_mgmt_len)
e_tbl_mgmt_name: "Name" "=" STRING_LITERAL
e_tbl_mgmt_desc: "Description" "=" STRING_LITERAL
e_tbl_mgmt_eqn: "Equation" "=" e_nested_str
e_tbl_mgmt_eqn_opts: e_tbl_mgmt_eqn_opts e_tbl_mgmt_eqn_opt | e_tbl_mgmt_eqn_opt
e_tbl_mgmt_eqn_opt: "@"e_tbl_mgmt_at_fns | "(" | ")" | "," | "==" | IDENTIFIER | e_nested_str | INT | DOUBLE_QUOTE STRING_LITERAL DOUBLE_QUOTE
e_tbl_mgmt_at_fns: "CompIf" | "Ctod" | "MID" | "Dtoc" | "Abs" | "Afternoon" | "Age"
e_tbl_mgmt_len: "Length" "=" INT
e_tbl_mgmt_type: "Type" "=" e_tbl_mgmt_types
e_tbl_mgmt_types: WI_CHAR_FIELD | WI_TIME_FIELD | WI_DATE_FIELD | WI_VIRT_CHAR | WI_VIRT_DATE
WI_CHAR_FIELD: "WI_CHAR_FIELD"
WI_TIME_FIELD: "WI_TIME_FIELD"
WI_DATE_FIELD: "WI_DATE_FIELD"
WI_VIRT_CHAR: "WI_VIRT_CHAR"
WI_VIRT_DATE: "WI_VIRT_DATE"
e_nested_str: DOUBLE_QUOTE e_tbl_mgmt_eqn_opts DOUBLE_QUOTE

e_tbl_mgmt_task_opts: "task" "." (d_add_field | d_rename_field | e_perf_task)
d_add_field: "AppendField" "field"
d_rename_field: "ReplaceField" STRING_LITERAL "," "field"
e_perf_task: "PerformTask"

e_tbl_mgmt_set_tsk_opts: "Set" "task" "=" (e_tbl_mgmt_st_nt | e_tbl_mgmt_inner)
e_tbl_mgmt_inner: "db" "." "TableManagement" e_tbl_mgmt_st e_tbl_opts 
e_tbl_mgmt_set_oth_nts: "Set" ("field" | "db") "=" "Nothing"
e_tbl_mgmt_st_nt: "Nothing"

---


e_tbl_opt: e_tbl_mgmt_opts e_tbl_mgmt_task_opt e_optionals
e_tbl_mgmt_opts: e_tbl_mgmt_opts e_tbl_mgmt_opt | e_tbl_mgmt_opt
e_tbl_mgmt_opt: "field" "." (e_tbl_mgmt_name | e_tbl_mgmt_desc | e_tbl_mgmt_type | e_tbl_mgmt_eqn | e_tbl_mgmt_len)
e_tbl_mgmt_name: "Name" "=" STRING_LITERAL
e_tbl_mgmt_desc: "Description" "=" STRING_LITERAL
e_tbl_mgmt_eqn: "Equation" "=" e_nested_str
e_tbl_mgmt_eqn_opts: e_tbl_mgmt_eqn_opts e_tbl_mgmt_eqn_opt | e_tbl_mgmt_eqn_opt
e_tbl_mgmt_eqn_opt: "@"e_tbl_mgmt_at_fns | "(" | ")" | "," | "==" | IDENTIFIER | e_nested_str | INT | DOUBLE_QUOTE STRING_LITERAL DOUBLE_QUOTE
e_tbl_mgmt_at_fns: "CompIf" | "Ctod" | "MID" | "Dtoc" | "Abs" | "Afternoon" | "Age"
e_tbl_mgmt_len: "Length" "=" INT
e_tbl_mgmt_type: "Type" "=" e_tbl_mgmt_types
e_nested_str: DOUBLE_QUOTE e_tbl_mgmt_eqn_opts DOUBLE_QUOTE

e_tbl_mgmt_types: WI_CHAR_FIELD | WI_TIME_FIELD | WI_DATE_FIELD | WI_VIRT_CHAR | WI_VIRT_DATE
WI_CHAR_FIELD: "WI_CHAR_FIELD"
WI_TIME_FIELD: "WI_TIME_FIELD"
WI_DATE_FIELD: "WI_DATE_FIELD"
WI_VIRT_CHAR: "WI_VIRT_CHAR"
WI_VIRT_DATE: "WI_VIRT_DATE"

e_tbl_mgmt_task_opt: "task" "." (d_add_field | d_rename_field)
d_add_field: "AppendField" "field"
d_rename_field: "ReplaceField" STRING_LITERAL "," "field"
e_optionals: e_tasks | st_nts |
perf_task: "." "PerformTask"
set_task: "=" "db" "." "TableManagement"
e_tasks: "task" (perf_task | set_task)