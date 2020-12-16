
class Community(object):

    # def __init__(self, name, proportion, time_befor, time_after):
    #     self.name = name
    #     self.proportion = proportion
    #     self.time_before = time_befor
    #     self.time_after = time_after

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def proportion(self):
        return self._proportion

    @proportion.setter
    def proportion(self, value):
        self._proportion = value

    @property
    def time_before(self):
        return self._time_before

    @time_before.setter
    def time_before(self, value):
        self._time_before = value

    @property
    def time_after(self):
        return self._time_after

    @time_after.setter
    def time_after(self, value):
        self._time_after = value