
class Util():


    def sortObjectCollection(self, theList):
        if theList.Contains('Id'):
            theList.Remove('Id')
            return ['Id'] + sorted(theList)
        return sorted(theList)

