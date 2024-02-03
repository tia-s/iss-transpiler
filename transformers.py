from lark import Transformer, v_args, Token, Tree

@v_args(inline=True)
class IDEATransformer(Transformer):   
    def __init__(self, translator):
        self.translator = translator
        self.var_list = []
        self.iss_list = []
        self.struct_list = []
        self.fns_list = []
        self.extract_list = []

    def token_str(self):
        return f"{self.value}"
    
    def tree_str(self):
        return f"{self.children}"
    
    Token.__repr__ = token_str
    Tree.__repr__ = tree_str
    
    """ Global Rules """
    def prog(self, start):
        return start

    def start(self, token):
        return token
    
    def iss_opts(self, *_):
        return {"iss_opts": self.iss_list}
    
    def iss_opt(self, token):
        self.iss_list.append(token)
        return token
    
    """ Set Functions """
    def i_st_opn_db(self, var_name, db_name):
        return {"var_name": var_name, "db_name": db_name}
    
    def i_st_nts(self, *tokens):
        return "set nothing"
    
    def i_st_nt(self, _):
        return "Nothing"
    
    """ Conditionals """
    def i_cnd_dcl(self, *tokens):
        return "if decl"
    
    def i_cnd_dcl_opts(self, *_):
        return
    
    def i_cnd_dcl_opt(self, *_):
        return 
    
    def i_cnd_hav(self, *_):
        return
    
    def i_cnd_els(self, *tokens):
        return "else call log"
    
    def i_cnd_tnl(self, *tokens):
        return "if end"
    
    """ Functions """
    def i_fns(self, *tokens):
        return {"fns": self.fns_list}
    
    def i_fn(self, *tokens):
        output = {"open_db": tokens[1], "fn": tokens[2]}
        self.fns_list.append(output)
        return output
    
    def i_st_mtd(self, *tokens):
        return tokens[-1]
    
    def i_st_fn(self, *tokens):
        return tokens[-1]

    def s_st_fns(self, token):
        return token
    
    def i_add_flds_to_inc(self, *tokens):
        return tokens
    
    def i_add_fld_to_inc(self, *tokens):
        return tokens[-1]
    
    def i_crt_virt_db(self, *tokens):
        return tokens
    
    """ Extraction Rules """
    def d_extract(self, *tokens):
        return tokens
        return {"extract": self.extract_list}
    
    def d_extract_opts(self, *tokens):
        # self.extract_list.append(tokens)
        return tokens
    
    def d_extract_opt(self, *tokens):
        return tokens
    
    def s_extract_opts(self, token):
        return token
    
    def e_add_extraction(self, *tokens):
        return tokens
    
    def e_perf_task_extract(self, *tokens):
        return tokens
    
    def e_add_extraction(self, *tokens):
        return tokens
    

    """ Struct Rules """
    def structs(self, *tokens):
        return {"structs": self.struct_list}    
    
    def struct(self, *tokens):
        return self.struct_list.append(tokens[1])
    
    def struct_opts(self, token):
        return token
    
    def struct_cond_decl(self, struct_name):
        return struct_name
    
    """ Sub Rules """
    def subs(self, *tokens):
        return tokens
    
    def sub_cond_decl(self, sub_name):
        return sub_name
    
    def subroutine_opts(self, *tokens):
        return tokens
    
    def subroutine_opt(self, token):
        return token
    
    def sub_ignore_warn(self, token):
        return token
    
    def sub_set(self, *tokens):
        return tokens[1]
    
    def sub_on_err(self, token):
        return token
    
    """ Var Rules """
    def vars_t(self, *_):
        return {"vars": self.var_list}

    def var(self, *tokens):
        output = {"type": tokens[0], "id": tokens[1], "op": tokens[2]}
        self.var_list.append(output)
        self.translator.declare_vars(output)
        return output
    
    def var_type_opts(self, token):
        return token
    
    def asgn_opts(self, token):
        return {"assign": token}

    def decl_opts(self, token):
        return {"declare": token}
    
    def type_opts(self, token):
        return token
    
    def val_opts(self, token):
        return token
    
    def var_type_decl_opts(self, token):
        return token
    
    def const_v(self):
        return "constant"

    def global_v(self):
        return "global"
    
    def reg_v(self):
        return "variable"
    
    def type_str(self):
        return "string"

    def type_obj(self):
        return "object"
    
    def type_double(self):
        return "double"
    

if __name__ == "__main__":
    pass