import getpass

UserGroupDir = {
  "matepeter": "BIM",
  "olivia": "BIM",
  "agnes.refi": "BIM",
  "pal.rada" : "BIM" ,
  "jozsef.szilagyi" : "BIM"
}



def getUserGroup():
    UserGroup = UserGroupDir.get(getpass.getuser())
    return(UserGroup)