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
        # account for having multiple stats to include
        # account for having add field to include (need to join back to summby)
        aggs = {"SM_SUM": ["sum"]}

        db_name = summ_dict["dbname"].replace('.IMD', '')
        fields_to_summarize = summ_dict["Add to Summarize"]
        agg_func = aggs[summ_dict["stat"]]
        fields_to_total = summ_dict["Add to Total"]

        self.indenter.write_to_file(f'wd.summBy("{db_name}", {fields_to_summarize}, agg_funcs={{key: {agg_func} if key != "{fields_to_summarize[0]}" else {agg_func + ["count"]} for key in {fields_to_total}}})')
        self.indenter.write_to_file(f'wd.renameCol(columns={{"{fields_to_summarize[0] + "_count"}": "NO_OF_RECS"}})')
        

        print(summ_dict)
        
class Indenter():
    def __init__(self):
        self.indent_level = 0

    def write_to_file(self, data, newlines=1):
        with open('output.py', 'a') as f:
            f.write(f"{data}" + "\n"*newlines + "\t"*self.indent_level)
