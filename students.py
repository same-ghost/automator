class Degree:


    def __init__(self, type):
        self.type = type
    # END Degree.__init__()


    def __str__(self):
        type = {'bm': 'Bachelor of Music', 'bs': 'Bachelor of Science', 'ba': 'Bachelor of Arts'}
        return f"{type[self.type]}"
    # END Degree.__str__()


    def audit(self):
        pass
    # END Degree.audit()


# END Degree()