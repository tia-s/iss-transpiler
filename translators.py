from abc import ABC, abstractmethod

class Translator(ABC):
    @abstractmethod
    def imports(self):
        pass


class RDMTranslator(Translator):
    def imports():
        pass