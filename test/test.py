import sys
s1 = "灰色; xxL"
s2 = "灰色; M"
s3 = "灰色; "
s4 = "灰色; S"

#@print(s.replace('XXL', '1'))


def size_to_num(x):
    x = str(x).upper()
    map = (('XXXS', '1'), ('XXXL', '9'),
           ('XXS', '2'), ('XXL', '8'),
           ('XS', '3'), ('XL', '7'),
           ('S', '4'), ('M', '5'), ('L', '6'))
    for t in map:
        s = t[0]
        num = t[1]
        if s in x:
            return x.replace(s, num)
    return x

print(size_to_num(s1))
print(size_to_num(s2))
print(size_to_num(s3))
print(size_to_num(s4))