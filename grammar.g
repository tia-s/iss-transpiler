%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.WS
%import common.CNAME -> IDENTIFIER

%ignore WS

start: (var_assgn | var_decl | T_COMMENT)*

T_COMMENT: /'[^\n]*/

// Variable Declaration
T_CONST: "CONST"
T_GLOBAL: "GLOBAL"
T_DIM: "DIM"

s_dcl: T_CONST | T_GLOBAL | T_DIM

T_AS: "AS"

// Type Keywords
T_STRING: "STRING"
T_OBJECT: "OBJECT"
T_DOUBLE: "DOUBLE"
T_PROJECT_MANAGEMENT: "PROJECTMANAGEMENT"
s_typ: T_STRING | T_OBJECT | T_DOUBLE | T_PROJECT_MANAGEMENT

s_bool: "FALSE" | "TRUE"
s_assgn_vals: s_bool | INT | STRING_LITERAL

// Operators
T_EQUAL: "="

// Assignments
var_assgn: s_dcl? IDENTIFIER T_EQUAL s_assgn_vals

// Declarations
var_decl: s_dcl IDENTIFIER T_AS s_typ