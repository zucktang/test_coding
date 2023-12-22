def generate_c(a, b):
    M = len(b)
    N = len(a)
    list_c = [0] * M
    set_a = set(a)
    
    for i in range(M):
        found = False 
        for x in range(N):
            for y in range(x, N):
               
                z = b[i] - a[x] - a[y]
                if z in set_a:
                    list_c[i] = 1
                    found = True
                    print(f"{a[x]} + {a[y]} + {z} = {b[i]}")
                    break
            if found:
                break 

    return list_c

a = [5, 2, 3, 4, 10, 11]
b = [5, 7, 26, 16, 5, 10, 12]
result_c = generate_c(a, b)
print(result_c)

time complexity = O(MN^2)