print("this is another test from api")

list1 = list(range(100))
print(list1)

list2 = [i for i in range(200)]

print(list2)

#for i in zip(list1,list2):
#    print(i)

tup1 = tuple(zip(list1,list2))

print(tup1)
