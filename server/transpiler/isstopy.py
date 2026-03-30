from lark import Lark
from translators import PandasTranslator
from transformers import IDEATransformer
from pathlib import Path


def main():
    DATA_DIR = Path(__file__).parent / "data"

    with open(DATA_DIR / 'grammar.g', 'r') as f, open(DATA_DIR / 'source.iss', 'r') as ff:
        grammar = f.read()
        text = ff.read()

    # contextual lexer needed since vb sometimes allows keywords to be used as identifiers (e.g. "Set" can be a keyword or an object name)
    parser = Lark(grammar=grammar, start='start', parser='lalr', lexer='contextual')

    tree = parser.parse(text)

    # translator = PandasTranslator()
    # transformer = IDEATransformer(translator)
    # flattened = transformer.transform(tree)

    print(tree.pretty())
    # print(flattened)

if __name__ == "__main__":
    main()


