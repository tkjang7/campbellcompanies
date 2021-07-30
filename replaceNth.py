def replacenth(instring,pattern,replacement,n=[1]):
    """

    Replace specified instance(s) of pattern in string.

      Positional arguments
        instring - input string
         pattern - regular expression pattern to search for
     replacement - replacement

      Keyword arguments
               n - list of instances requested to be replaced [default [1]]

    """

    import re
    outstring=''
    i=0
    for j,m in enumerate(re.finditer(pattern,instring)):
        if j+1 in n: outstring+=instring[i:m.start()]+replacement
        else: outstring+=instring[i:m.end()]
        i=m.end()
    outstring+=instring[i:]
    return outstring
    