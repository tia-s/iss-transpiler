from abc import ABC, abstractmethod

class Translator(ABC):
    @abstractmethod
    def imports(self):
        pass


class PandasTranslator(Translator):
    def imports():
        pass