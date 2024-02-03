from abc import ABC, abstractmethod

class Translator(ABC):
    @abstractmethod
    def imports(self):
        pass

    @abstractmethod
    def declare_vars(self):
        pass


class RDMTranslator(Translator):
    def imports(self):
        pass

    def declare_vars(self, var_dict):
        var_type = var_dict["type"]
        var_name = var_dict["id"]
        var_op = var_dict["op"]

        with open("output.py", "w+") as f:
            f.write(f"{var_type, var_name, var_op}")