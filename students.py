class Students:

    def __init__(self, id, degree):
        self.id = id
        self.degree = None
        self.core = []
    # END Students.__init__()


    def ger(self, course, term, grade, ger):
        if course in self.core:
            break
        else:
            self.core.append(course)
            
    # END Students.ger()