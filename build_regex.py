dict_regex = {}
keys = range(100)
regex_string = 're.compile(r"^'
for i in keys:
    dict_regex[i] = regex_string + str(i) + '")'
print(dict_regex)
print(dict_regex[10])
