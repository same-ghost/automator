class Degree:


    def __init__(self, type):
        self.type = type
        self.gers = ['FYW', 'WR', 'NWL', 'NWL', 'HB1', 'HB2', 'TA', 'HA', 'VP',
        'MR', 'FL', 'UQ', 'MB', 'NE', 'WC']

        if self.type == 'bm':
            self.gers.pop(3)
            self.gers.pop(8)
        elif self.type == 'ba':
            self.gers[3] = 'NW'
    # END Degree.__init__()


    def __str__(self):
        type = {'bm': 'Bachelor of Music', 'bs': 'Bachelor of Science', 'ba': 'Bachelor of Arts'}
        return f"{type[self.type]}"
    # END Degree.__str__()


# END Degree()


class Student:


    def __init__(self, id='0', name='', degree='ba'):
        self.id = id
        self.name = name
        self.degree = Degree(degree)
        self.completed = []
    # END Student.__init__()


    def __str__(self):
        return f"{self.id} {self.name} {self.degree}"
    # END Student.__str__()


    def add_ger(self, course, term, grade, ger):
        pass
# END Student()