def get_city(o):
    return o.get('city')


def stable_sort(arr, key):
    result = []
    for _, item in sorted(enumerate(arr), key=lambda x: (key(x[1]), x[0])):
        result.append(item)
    return result

arr = [
    {'name': 'cat', 'city': 'BKK'},
    {'name': 'fish', 'city': 'BKK'},
    {'name': 'ant', 'city': 'BKK'},
    {'name': 'dog', 'city': 'LONDON'}
]

arr_stable_sorted = stable_sort(arr, key=get_city)
print(arr_stable_sorted)