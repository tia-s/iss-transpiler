from lark import Transformer, v_args, Token

@v_args(inline=True)
class IDEATransformer(Transformer):   
    def token_str(self):
        return f"{self.value}"
    
    Token.__repr__ = token_str