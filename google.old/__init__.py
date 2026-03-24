# This is a stub file to make Pylance happy in a local environment.
# The actual google.colab library only exists in the Colab runtime.

class _Drive:
    def mount(self, path):
        pass

class _UserData:
    def get(self, key):
        return None

drive = _Drive()
userdata = _UserData()
