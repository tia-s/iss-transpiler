from lark import Lark
from translators import RDMTranslator
from transformers import IDEATransformer

def main():
    with open('rules.g', 'r') as f, open('source.iss', 'r') as ff:
        grammar = f.read()
        text = ff.read()

    parser = Lark(grammar=grammar, start='start', parser='lalr')

    tree = parser.parse(text.upper())

    # translator = RDMTranslator()
    # transformer = IDEATransformer(translator)
    # flattened = transformer.transform(tree)

    print(tree.pretty())

if __name__ == "__main__":
    main()



