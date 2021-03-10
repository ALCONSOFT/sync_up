variable_name = "I am a variable"
print("This string has a variable coming up next: %s" % variable_name)
var_1 = 'HOla'
var_2 = 'Esta'
var_3 = 'PRueba'
var = var_1 + " " + var_2 + " some static text " + var_3 + "!"
print(var)

print("%s %s some static text %s!"%(var_1,var_2,var_3))

field_name = 'campo'
value_to_select = 1
sql_query = "\"%s\" = '%d'" % (field_name,value_to_select)
print(sql_query)