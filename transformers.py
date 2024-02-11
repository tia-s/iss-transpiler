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
    
    # Token.__repr__ = token_str
    # Tree.__repr__ = tree_str
    
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
    

@v_args(inline=True)
class NewTransformer(Transformer):   
    # consider creating instances of objects instead of making dictionaries/lists like this
    def __init__(self, translator):
        self.translator = translator
        self.add_fields_to_total = []
        self.add_fields_to_summarize = []
        self.add_fields_to_inc = []
        self.summarize_task_opts = {}
        self.have_recs_list = []
        self.summarize_agg_funcs_list = []

        Token.__repr__ = lambda self: self.value

    def _reset(self):
        self.add_fields_to_total = []
        self.add_fields_to_summarize = []
        self.add_fields_to_inc = []
        self.summarize_task_opts = {}
        self.have_recs_list = []
        self.summarize_agg_funcs_list = []


    """ Global Rules """
    def STRING_LITERAL(self, token):
        return str(token[1:-1])

    def COMMENT(self, comment):
        comment = comment.replace('\'', '')
        self.translator.comment(comment)
    
    def s_bools(self, token):
        return token
    
    """ Struct Rules """
    def struct_cond_decl(self, _, id):
        self.translator.define_function(id)

    def struct_cond_tnl(self, *_):
        self.translator.end_function()
    
    def st_nts(self, *_):
        return "Set Nothings"
    
    def iss_match_method(self, token):
        return token
    
    """ Bp Function Rules """
    def bp_std_fns(self, *_):
        # could either set these to be empty here or empty them after call to function
        self._reset()

    def have_records_check_decl(self, *_):
        # call if have records
        return "Have Records Check"
    
    def have_records_opts(self, *_):
        self.translator.bp_cond_check(self.have_recs_list)
    
    def have_records_opt(self, token):
        self.have_recs_list.append(token)

    def have_records(self, _, id):
        return id

    def have_records_check_tnl(self, *_):
        self.translator.bp_cond_end()
    
    def have_records_check_else(self, _, bp_method_opts):
        # log to file
        return {"log": bp_method_opts}
    
    def bp_method_opts(self, id):
        return id
    
    def st_open_db(self, id):
        self.translator.open_table(id)
    
    """ Summarization Rules """
    def d_summarize(self, *_):
        # summby
        self.translator.summarize(self.summarize_task_opts)
    
    def e_summarize_opts(self, *_):
        return self.summarize_task_opts
    
    def e_summarize_opt(self, token):
        return token
    
    def e_summarize_task_opts(self, token):
        return token

    def s_summarize_task_opts(self, token):
        self.summarize_task_opts.update(token)

    def e_add_fields_to_summarize(self, *_):
        return {"Add to Summarize": self.add_fields_to_summarize}
    
    def e_add_field_to_summarize(self, token):
        self.add_fields_to_summarize.append(token)

    def e_add_fields_to_total(self, *_):
        return {"Add to Total": self.add_fields_to_total}
    
    def e_add_field_to_total(self, token):
        self.add_fields_to_total.append(token)

    def e_add_fields_to_inc(self, *_):
        return {"Add to Inc": self.add_fields_to_inc}
    
    def e_add_field_to_inc(self, token):
        self.add_fields_to_inc.append(token)

    def e_summ_criteria(self, *tokens):
        # deal with criteria (call filter)
        # return {"Criteria": tokens[1]}
        return {"Criteria": ""}

    def e_summ_output_db_name(self, *_):
        return {"Output DB Name": ""}

    def e_summ_create_percent_field(self, token):
        return {"create_percnt": token}
    
    def e_summ_statistics_to_include(self, token):
        return token

    def e_perform_task(self, *_):
        return {"Perform Task": ""}

    def e_summ_db_name(self, id):
        self.summarize_task_opts.update({"dbname": id})
    
    def s_stats_opts(self, *_):
        return {"stats": self.summarize_agg_funcs_list}
    
    def SM_SUM(self, *_):
        self.summarize_agg_funcs_list.append("SM_SUM")
        
    def SM_AVERAGE(self, *_):
        self.summarize_agg_funcs_list.append("SM_AVERAGE")







    # def std_fns_decl(self, *tokens):
    #     return {"open_db": tokens[0], "fn": tokens[1]}
    
    # def std_fns_opts(self, *tokens):
    #     return tokens
    
    # def st_fn(self, token):
    #     return token
    



    


    
if __name__ == "__main__":
    pass