from lark import Transformer, v_args, Token, Tree

"""
account for if within join (bp with if after if)
"""

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

        self.join_task_opts = {}
        self.add_s_fields_to_inc = []
        self.add_p_fields_to_inc = []

        self.extract_task_opts = {}
        self.extract_add_fields_to_inc = []

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


    """ Join Rules """
    def d_join(self, *_):
        self.translator.join(self.join_task_opts)

    def e_join_opt(self, token):
        return token

    def e_join_task_opts(self, token):
        return token

    def s_join_task_opts(self, token):
        self.join_task_opts.update(token)

    def e_file_to_join(self, id):
        return {"file_to_join": id}
    
    def e_include_all_p_fields(self, *_):
        return {"p_fields": "all"}
    
    def e_include_all_s_fields(self, *_):
        return {"s_fields": "all"}
    
    def e_add_s_fields_to_inc(self, *_):
        return {"s_fields": self.add_s_fields_to_inc}
    
    def e_add_s_field_to_inc(self, field):
        self.add_s_fields_to_inc.append(field)

    def e_add_p_fields_to_inc(self, *_):
        return {"p_fields": self.add_p_fields_to_inc}

    def e_add_p_field_to_inc(self, field):
        self.add_p_fields_to_inc.append(field)

    def e_add_match_key(self, s1, s2, a):
        return {"match_keys": [s1, s2, a]}
    
    def e_join_create_virt_database(self, token):
        return {"create_virtual_db": token}
    
    def e_join_perform_task(self, *tokens):
        return {"perform_task": tokens[-1]}
    
    def s_join_types(self, token):
        return token
    
    def WI_JOIN_ALL_IN_PRIM(self, *_):
        return "WI_JOIN_ALL_IN_PRIM"
    
    def WI_JOIN_MATCH_ONLY(self, *_):
        return "WI_JOIN_MATCH_ONLY"
    
    def e_join_db_name(self, id):
        self.join_task_opts.update({"db_name": id})


    """ Extract Rules """
    def d_extract(self, *_):
        self.translator.extract(self.extract_task_opts)

    def e_extract_opt(self, token):
        return token
    
    def e_extract_task_opts(self, token):
        return token
    
    def s_extract_task_opts(self, token):
        self.extract_task_opts.update(token)

    def e_extract_include_all_fields(self, *_):
        return {"fields": "all"}
    
    def e_extract_add_fields_to_inc(self, *_):
        return {"fields": self.extract_add_fields_to_inc}
    
    def e_extract_add_field_to_inc(self, token):
        self.extract_add_fields_to_inc.append(token)

    def e_add_extraction(self, *tokens):
        return {"filter": ""}
    
    def e_extract_create_virt_database(self, token):
        return {"create_virtual_database": token}
    
    def e_extract_perform_task(self, *_):
        return {"perform_task": ""}
    
    def e_extract_db_name(self, id):
        self.extract_task_opts.update({"db_name": id})



    # def std_fns_decl(self, *tokens):
    #     return {"open_db": tokens[0], "fn": tokens[1]}
    
    # def std_fns_opts(self, *tokens):
    #     return tokens
    
    # def st_fn(self, token):
    #     return token
    



    


    
if __name__ == "__main__":
    pass