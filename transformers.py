from lark import Transformer, v_args, Token, Tree

"""
account for if within join (bp with if after if)
account for set nothings in add/rename cols
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
        self.extract_add_fields = []

        self.export_task_opts = {}
        self.export_add_fields_to_inc = []

        self.cleanup_task_opts = {}
        self.cleanup_delete_files = []

        self.table_manage_task_opts = {}
        self.table_manage_dicts = []

        self.visual_connect_task_opts = {}
        self.visual_connect_add_fields_to_include = []
        self.visual_connect_add_assgn = []
        self.visual_connect_add_relations = []

        self.dup_key_exclude_task_opts = {}

        self.dup_key_detect_task_opts = {}
        self.dup_key_add_fields = []

        self.sort_task_opts = {}
        self.sort_add_keys = []

        self.index_task_opts = {}

        self.top_recs_extract_opts = {}
        self.top_recs_keys = []
        self.top_recs_fields = []

        self.append_db_task_opts = {}
        self.append_db_add_dbs = []

        self.import_odbc_task_opts = {}
        self.import_odbc_file_opts = []

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
        # self.translator.comment(comment)
    
    def s_bools(self, token):
        return token
    
    def e_top_recs_extract_db_name(self, token):
        self.top_recs_extract_opts.update({"db_name": token})

    def s_field_wi_types(self, token):
        return token
    
    def WI_BOOL(self, *_):
        return "WI_BOOL"
    
    def STRING_LITERAL(self, token):
        return token.replace(".IMD", "")
    
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
    
    def e_summ_use_field_from_first_occurrence(self, token):
        return {"use_field_from_first": token}
    
    def e_summ_result_name(self, token):
        return {"result_name": token}
    
    def e_summ_use_quick_summarization(self, token):
        return {"use_quick_summ": token}
    
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

    def SM_MIN(self, *_):
        self.summarize_agg_funcs_list.append("SM_MIN")

    def SM_MAX(self, *_):
        self.summarize_agg_funcs_list.append("SM_MAX")

    def SM_VARIANCE(self, *_):
        self.summarize_agg_funcs_list.append("SM_VARIANCE")

    def SM_STD_DEV(self, *_):
        self.summarize_agg_funcs_list.append("SM_STD_DEV")


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
    
    def WI_JOIN_NOC_SEC_MATCH(self, *_):
        return "WI_JOIN_NOC_SEC_MATCH"
    
    def WI_JOIN_ALL_REC(self, *_):
        return "WI_JOIN_ALL_REC"
    
    def WI_JOIN_NOC_PRI_MATCH(self, *_):
        return "WI_JOIN_NOC_PRI_MATCH"
    
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

    def e_add_fields(self, *_):
        return {"created_fields": self.extract_add_fields}
    
    def e_add_field(self, *tokens):
        self.extract_add_fields.append(tokens)

    def e_add_extraction(self, *tokens):
        return {"filter": ""}
    
    def e_extract_create_virt_database(self, token):
        return {"create_virtual_database": token}
    
    def e_extract_perform_task(self, *_):
        return {"perform_task": ""}
    
    def e_extract_db_name(self, id):
        self.extract_task_opts.update({"db_name": id})

    """ Export Rules """
    def d_export(self, *_):
        self.translator.export(self.export_task_opts)

    def e_export_opt(self, token):
        return token
    
    def e_export_task_opts(self, token):
        return token
    
    def s_export_task_opts(self, token):
        self.export_task_opts.update(token)

    def e_export_add_fields_to_inc(self, *_):
        return {"fields": self.export_add_fields_to_inc}

    def e_export_add_field_to_inc(self, field):
        self.export_add_fields_to_inc.append(field)

    def e_export_perform_task(self, *_):
        return {"perform_task": ""}
    
    def e_export_eqn(self, token):
        return {"eqn": token}
    

    """ Cleanup Rules """
    def d_cleanup(self, *_):
        self.translator.cleanup(self.cleanup_task_opts)

    def e_cleanup_task_opts(self, token):
        self.cleanup_task_opts.update(token)

    def e_cleanup_task_opt(self, token):
        return token
    
    def e_cleanup_delete_files(self, *_):
        return {"files": self.cleanup_delete_files}
    
    def e_cleanup_delete_file(self, id):
        self.cleanup_delete_files.append(id)

    """ Table Management Rules """
    def e_tbl_mgmt_opt(self, token):
        return token
    
    def e_tbl_mgmt_field_opts(self, token):
        return token
    
    def e_tbl_mgmt_task_opts(self, token):
        return token
    
    def s_tbl_mgmt_field_opts(self, token):
        self.table_manage_task_opts.update(token)
    
    def e_tbl_mgmt_name(self, token):
        return {"name": token}
    
    def e_tbl_mgmt_desc(self, token):
        return {"description": token}

    def e_tbl_mgmt_len(self, token):
        return {"length": token}
    
    def s_tbl_mgmt_types(self, token):
        return token
        
    def e_tbl_mgmt_type(self, token):
        return {"field_type": token}
    
    def e_tbl_mgmt_decimals(self, token):
        return {"decimals": token}
    
    def e_tbl_mgmt_eqn(self, token):
        return {"equation": token}
    
    def e_tbl_mgmt_append_field(self, *_):
        self.translator.table_manage(self.table_manage_task_opts)
        self.table_manage_task_opts = {}
    
    
    def e_tbl_mgmt_replace_field(self, token):
        self.table_manage_task_opts.update({"replace": token})
        self.translator.table_manage(self.table_manage_task_opts)
        self.table_manage_task_opts = {}

    
    # def WI_VIRT_CHAR(self, _):
    #     return "WI_VIRT_CHAR"

    """ Visual Connect Rules """
    def d_visual_connect(self, *_):
        self.translator.connect(self.visual_connect_task_opts)

    def e_visual_connect_opt(self, token):
        return token
    
    def e_visual_connect_task_opts(self, token):
        return token
    
    def s_visual_connect_task_opts(self, token):
        self.visual_connect_task_opts.update(token)

    def e_visual_connect_add_fields_to_include(self, *_):
        return {"fields_to_include": self.visual_connect_add_fields_to_include}

    def e_visual_connect_add_field_to_include(self, db, field):
        self.visual_connect_add_fields_to_include.append((db, field))

    def e_visual_connect_add_relations(self, *_):
        return {"add_relation": self.visual_connect_add_relations}
    
    def e_visual_connect_add_relation(self, token):
        return token
    
    def s_visual_connect_add_relation_opts(self, *tokens):
        return {"add_relation": "do later"}
    
    def e_visual_connect_master_database(self, token):
        return {"master_db": token}
    
    def e_visual_connect_append_database_names(self, token):
        return {"append_db_names": token}
    
    def e_visual_conenct_include_all_primary_recs(self, token):
        return {"include_all_prim_recs": token}
    
    def e_visual_connect_add_database(self, token):
        return {"add_db": token}
    
    def e_visual_connect_create_virt_database(self, token):
        return {"create_virt_db": token}

    def e_visual_connect_output_db_name(self, token):
        return {"output_db": token}
    
    def e_visual_connect_perf_task(self, *_):
        return {"perf_task": ""}
    
    def e_visual_connect_db_name(self, token):
        self.visual_connect_task_opts.update({"db_name": token})

    def e_visual_connect_add_assgns(self, *_):
        self.visual_connect_task_opts.update({"add_assigns": self.visual_connect_add_assgn})

    def e_visual_connect_add_assgn(self, *tokens):
        self.visual_connect_add_assgn.append((tokens[0], tokens[1]))

    """ Duplicate Key Exclusion Rules """
    def d_dup_key_exclude(self, *_):
        self.translator.dup_key_exclude(self.dup_key_exclude_task_opts)

    def e_dup_key_exclude_opt(self, token):
        return token
    
    def e_dup_key_exclude_task_opts(self, token):
        return token
    
    def s_dup_key_exclude_task_opts(self, token):
        self.dup_key_exclude_task_opts.update(token)

    def e_dup_key_include_all_fields(self, *_):
        return {"fields": "all"}
    
    def e_dup_key_add_key(self, *tokens):
        return {"key": tokens[0]}
    
    def e_dup_key_different_field(self, token):
        return {"diff_field": token}
    
    def e_dup_key_create_virt_database(self, token):
        return {"virt_db": token}
    
    def e_dup_key_perf_task(self, *_):
        return {"perf_task": ""}
    
    def e_dup_key_exclude_db_name(self, token):
        self.dup_key_exclude_task_opts.update({"db_name": token})

    """ Duplicate Key Detection Rules """
    def d_dup_key_detect(self, *_):
        self.translator.dup_key_detect(self.dup_key_detect_task_opts)
    
    def e_dup_key_detect_opt(self, token):
        return token
    
    def e_dup_key_detect_task_opts(self, token):
        return token
    
    def s_dup_key_detect_task_opts(self, token):
        self.dup_key_detect_task_opts.update(token)

    def e_dup_key_detect_add_fields_to_inc(self, *_):
        return {"fields": self.dup_key_add_fields}
    
    def e_dup_key_detect_add_field_to_inc(self, token):
        self.dup_key_add_fields.append(token)

    def e_dup_key_detect_add_key(self, *tokens):
        return {"key": (tokens[0], tokens[1])}
    
    def e_dup_key_detect_output_duplicates(self, token):
        return {"output_duplicates": token}
    
    def e_dup_key_detect_create_virt_database(self, token):
        return {"create_virt_db": token}
    
    def e_dup_key_detect_perf_task(self, *_):
        return {"perf_task": ""}
    
    def e_dup_key_detect_db_name(self, token):
        self.dup_key_detect_task_opts.update({"db_name": token})


    """ Sort Rules """
    def d_sort(self, *_):
        self.translator.sort(self.sort_task_opts)

    def e_sort_opt(self, token):
        return token
    
    def e_sort_task_opts(self, token):
        return token
    
    def s_sort_task_opts(self, token):
        self.sort_task_opts.update(token)

    def e_sort_add_keys(self, *_):
        return {"keys": self.sort_add_keys}
    
    def e_sort_add_key(self, *tokens):
        self.sort_add_keys.append((tokens[0], tokens[1]))

    def e_sort_perf_task(self, *_):
        return {"perf_task": ""}
    
    def e_sort_db_name(self, token):
        self.sort_task_opts.update({"db_name": token})

    """ Index Rules """
    def d_index(self, *_):
        self.translator.index(self.index_task_opts)

    def e_index_opt(self, token):
        return token

    def s_index_task_opts(self, token):
        self.index_task_opts.update(token)

    def e_index_add_key(self, *tokens):
        return {"key": (tokens[0], tokens[1])}
    
    def e_index_index(self, token):
        return {"index": token}
    
    """ Top Records Extraction Rules """
    def d_top_recs_extract(self, *_):
        self.translator.top_recs_extract(self.top_recs_extract_opts)

    def e_top_recs_extract_opt(self, token):
        return token
    
    def e_top_recs_extract_task_opts(self, token):
        return token
    
    def s_top_recs_extract_task_opts(self, token):
        self.top_recs_extract_opts.update(token)

    def e_top_recs_extract_add_fields_to_inc(self, *_):

        return {"fields": self.top_recs_fields}
    
    def e_top_recs_extract_add_field_to_inc(self, token):
        self.top_recs_fields.append(token)
    
    def e_top_recs_extract_add_keys(self, *_):
        return {"keys": self.top_recs_keys}
    
    def e_top_recs_extract_add_key(self, *tokens):
        self.top_recs_keys.append((tokens[0], tokens[1]))

    def e_top_recs_extract_output_file(self, token):
        return {"output_file": token}
    
    def e_top_recs_extract_recs_to_extract(self, token):
        return {"no_of_recs": token}
    
    def e_top_recs_extract_create_virt_db(self, token):
        return {"virt_db": token}

    def e_top_recs_extract_perf_task(self, *_):
        return {"perf_task": ""}

    
    """ Append Database Rules """
    def d_append_db(self, *_):
        self.translator.append_db(self.append_db_task_opts)

    def e_append_db_opt(self, token):
        return token
    
    def e_append_db_task_opts(self, token):
        return token
    
    def s_append_db_task_opts(self, token):
        self.append_db_task_opts.update(token)

    def e_append_db_add_databases(self, *_):
        return {"dbs": self.append_db_add_dbs}
    
    def e_append_db_add_database(self, token):
        self.append_db_add_dbs.append(token)

    def e_append_db_perf_task(self, *_):
        return {"perf_task": ""}

    def e_append_db_db_name(self, token):
        self.append_db_task_opts.update({"db_name": token})

    """ Import ODBC Rules"""
    def d_import(self, *_):
        self.translator.import_odbc(self.import_odbc_task_opts)

    def e_import_opt(self, token):
        self.import_odbc_task_opts.update(token)
    
    def e_import_client_opts(self, token):
        return token
    
    def e_import_db_name(self, token):
        return {"db_name": token}

    def e_import_odbc_file(self, *_):
        return {"import_odbc": self.import_odbc_file_opts}

    def e_import_char_len(self, token):
        return ("chr", token)
    
    def s_import_odbc_file_opts(self, *_):
        return self.import_odbc_file_opts
        
    def s_import_odbc_file_opt(self, token=''):
        if token and token != '"':
            self.import_odbc_file_opts.append(token)

    def e_count_db_records(self, *_):
        return {"db_records": ""}
    
    def e_client_close_all(self, *_):
        return {"close_all": ""}


    # def std_fns_decl(self, *tokens):
    #     return {"open_db": tokens[0], "fn": tokens[1]}
    
    # def std_fns_opts(self, *tokens):
    #     return tokens
    
    # def st_fn(self, token):
    #     return token

    
if __name__ == "__main__":
    pass