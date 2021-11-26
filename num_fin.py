def is_float(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

def num_fin(string):
    place_n1 = 0
    place_n2 = string.find(' ')-1
    while place_n2!=-1:
        place_n2 = place_n2 + place_n1 + 1
        con=string[place_n1:place_n2]
        if is_float(con) or con==' ':
            string = string[:place_n1]+string[place_n2+1:]
        else:
            place_n1=place_n2+place_n1+1
        place_n2=string[place_n1+1:].find(' ')
        # place_n2=place_n2+place_n1+1
        # print(string)

    place_n2=len(string)
    con = string[place_n1:place_n2]
    if is_float(con):
        string = string[:place_n1] + string[place_n2:]
    # string = string.replace('  ', ' ')
    string = string.replace('  ', '')
    # print(string)
    return string