%import common.ESCAPED_STRING -> STRING_LITERAL
%import common.INT -> INT
%import common.WS
%import common.CNAME -> IDENTIFIER

%ignore WS

start: (stmt_var_assgn | stmt_var_decl | T_COMMENT)*

T_COMMENT: /'[^\n]*/

// Arithmetic Operators
T_OP_EXP: "^"
T_OP_MUL: "*"
T_OP_DIV: "/"
T_OP_INT_DIV: "\"
T_OP_MOD: "MOD"
T_OP_ADD: "+"
T_OP_SUB: "-"
T_OP_CONCAT: "&"

// Bit Shift Operators
T_LSHIFT: "<<"
T_RSHIFT: ">>"

// Assignment Operators
T_OP_ASSGN: "="
T_OP_EXP_ASSGN: "^="
T_OP_MUL_ASSGN: "*="
T_OP_DIV_ASSGN: "/="
T_OP_INT_DIV_ASSGN: "\="
T_OP_ADD_ASSGN: "+="
T_OP_SUB_ASSGN: "-="
T_OP_LSHIFT_ASSGN: "<<="
T_OP_RSHIFT_ASSGN: ">>="
T_OP_CONCAT_ASSGN: "&="

// Comparison Operators
T_OP_LT: "<"
T_OP_LTE: "<="
T_OP_GT: ">"
T_OP_GTE: ">="
T_OP_NEQ: "<>"
T_OP_IS: "Is"
T_OP_IS_NOT: "ISNOT"
T_OP_LIKE: "LIKE"

// Bitwise Operators
T_OP_AND: "AND"
T_OP_NOT: "NOT"
T_OP_OR: "OR"
T_OP_XOR: "XOR"
T_OP_AND_ALS: "ANDALSO"
T_OR_ELS: "ORELSE"
T_IS_FALSE: "ISFALSE"
T_IS_TRUE: "ISTRUE"
