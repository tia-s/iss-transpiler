from abc import ABC, abstractmethod

class Translator(ABC):
    @abstractmethod
    def imports(self):
        pass

    @abstractmethod
    def comment(self, comment):
        pass

    @abstractmethod
    def declare_vars(self):
        pass

# make each translator have a write to file and whenever you call a function make it write to the file

class RDMTranslator(Translator):
    # fix indentation on functions and bp fns (whenever i have two of them, second indentation isnt right)
    # fix quotes (make all single/double)
    # fix new line after comments
    def __init__(self):
        with open('output.py', 'w') as f:
            f.write('from DataAnalytics import DataAnalytics\n')
            f.write('wd = DataAnalytics()\n\n')
        self.indenter = Indenter()

    def comment(self, comment):
        self.indenter.write_to_file(f'#{comment}')

    def define_function(self, id):
        self.indenter.indent_level += 1
        self.indenter.write_to_file(f'def {id}():')

    def end_function(self):
        self.indenter.indent_level -= 1

    def bp_cond_check(self, lst):
        # account for the "and"s
        for x in lst:
            x = x.replace('.IMD', '')
            self.indenter.indent_level += 1
            self.indenter.write_to_file(f'if not wd.open("{x}").empty:')

    def bp_cond_end(self):
        self.indenter.indent_level -= 1

    def imports(self):
        pass

    def open_table(self, id):
        id = id.replace('.IMD', '')
        self.indenter.write_to_file(f'wd.open("{id}")')

    def declare_vars(self, var_dict):
        var_type = var_dict['type']
        var_name = var_dict['id']
        var_op = var_dict['op']

        self.indenter.write_to_file(f'{var_type, var_name, var_op}')

    def summarize(self, summ_dict):
        # could be criteria, could have fields to inc so could have no criteria & no inc, have criteria but no inc, have inc but not criteria or have no criteria or inc
        # account for having add field to include (need to join back to summby)
        aggs = {"SM_SUM": ["sum"], "SM_AVERAGE": ["mean"]}

        db_name = summ_dict["dbname"].replace('.IMD', '')
        fields_to_summarize = summ_dict["Add to Summarize"]
        agg_func = [aggs.get(stat, "") for stat in summ_dict["stats"]]
        agg_func = [stat[0] for stat in agg_func if stat]
        fields_to_total = summ_dict["Add to Total"] 
        count_dict = {fields_to_summarize[0]: ["count"] if fields_to_summarize[0] not in fields_to_total else agg_func + ["count"]}
        fields_to_total = fields_to_total + [f"{fields_to_summarize[0]}"] if fields_to_summarize[0] not in fields_to_total else fields_to_total
        criteria = summ_dict["Criteria"]

        self.indenter.write_to_file(f'wd.summBy("{db_name}", {fields_to_summarize}, agg_funcs={{key: {agg_func} if key != "{fields_to_summarize[0]}" else {count_dict[fields_to_summarize[0]]} for key in {fields_to_total}}})')
        self.indenter.write_to_file(f'wd.renameCol(columns={{"{fields_to_summarize[0] + "_count"}": "NO_OF_RECS"}})')
        
        if criteria:
            self.indenter.write_to_file(f'wd.extract("{db_name}", filter="{criteria}")')
            print("criteriaa")

        if "Add to Inc" in summ_dict:
            fields_to_inc = summ_dict["Add to Inc"]
            self.indenter.write_to_file(f'wd.join("{db_name}", right=wd.db("{db_name + "_summ"}"){[fields_to_inc]}, how="left")')

        print(summ_dict)
        print(agg_func)
        
class Indenter():
    def __init__(self):
        self.indent_level = 0

    def write_to_file(self, data, newlines=1):
        with open('output.py', 'a') as f:
            f.write(f"{data}" + "\n"*newlines + "\t"*self.indent_level)
