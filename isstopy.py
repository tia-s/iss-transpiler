from lark import Lark
from translators import RDMTranslator
from transformers import NewTransformer

def main():
    with open('n_grammar.g', 'r') as f, open('source.iss', 'r') as ff:
        grammar = f.read()
        text = ff.read()

    parser = Lark(grammar=grammar, start='prog', parser='lalr')

    tree = parser.parse(text)

    translator = RDMTranslator()
    transformer = NewTransformer(translator)
    flattened = transformer.transform(tree)

    print(tree.pretty())
    print(flattened)

if __name__ == "__main__":
    main()



