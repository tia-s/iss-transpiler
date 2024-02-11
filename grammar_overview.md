A. Terminals defined at top. ISS is case-insensitive so must account for different cases.
B. prog is augmented start symbol
C. start can contain any of iss_opts (possible options in iss file)
D. iss_opts can be repeated (so one or more iss_opt (which are the actual options))
E. iss_opt can be a struct (function) or a COMMENT
    - this allows comments to be outside of any enclosing struct
F. struct starts with "Function" IDENTIFIER and ends with END FUNCTION.
G. specify the types of functions inside struct_opts
H. i_fns inside struct can be a bp_std_fn (bulletproof), std_fn (standard functions - have opendatabase etc) or non-standard (without the opendatabase etc)

Standard Functions
1. Can start with either of:
- If have records (including "And")
- Open database

might need rule for i_fns and i_fn to allow multiple i_fn in i_fns
might have issue if bp is in a non-standard function (would have to separate out the bp_decl)

for now using specific names (field, db and task) since it's too hard to make unambiguous with IDENTIFIER

Prob dont need case (KISS)::
SET: "Set" | "SET" | "set"
CLIENT: "Client" | "CLIENT" | "client"
OPENDATABASE: "OpenDatabase" | "openDatabase" | "OPENDATABASE" | "opendatabase"
NORESULTSLOG: "NORESULTSLOG" | "noresultslog" | "NoResultsLog"
TABLEDEF: "TableDef" | "TABLEDEF" | "tableDef" | "tabledef"
NEWFIELD: "NewField" | "NEWFIELD" | "newField" | "newfield"

not sure why i needed to specify @ in equation if its a string literal\

---

think i should put st_nothing as optional after std function instead of as a separate type of std_opt?

do i put it inside each std_fns_opts?
make task.append/rename field just one then task.performtask + set nothings