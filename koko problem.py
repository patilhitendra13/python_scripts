from math import ceil
from math import gcd
from copy import deepcopy

def findspeed(piles,h):
   l = 1
   r = max(piles)

   while(l<r):
        mid = (l+r)//2
        time = 0
        for p in piles:
            time+=ceil(p/mid)
        
        if time>h:
            l=mid+1
        else:
            r=mid
   return l
    
            

print(findspeed(piles = [2,2], h = 2))
