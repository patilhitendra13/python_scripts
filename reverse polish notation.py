class Solution(object):
    def evalRPN(self, tokens):
        """
        :type tokens: List[str]
        :rtype: int
        """

        res = []

        for char in tokens:
            if char.isdigit() or (char[0] == '-' and char[1:].isdigit()):
                res.append(char)
            else:
                sec = res.pop()
                fir = res.pop()
                res.append(int(eval(str(fir)+char+str(sec))))
        
        assert len(res)==1

        return res[0]
        

inp = ["10","6","9","3","+","-11","*","/","*","17","+","5","+"]

print(Solution.evalRPN(Solution,inp))