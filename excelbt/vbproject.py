import os

class Module(object):
    extension = '.bas'

    def __init__(self, name, code):
        self.name = name
        self.code = code

    def export(self, path):
        destination = os.path.join(path, self.filename)
        file(destination, 'w+').write(self.code)
        return self.filename

    @property
    def filename(self):
        return self.name + self.extension

class ClassModule(Module):
    extension = '.cls'

class VBProject(object):
    def __init__(self, modules=None):
        self.modules = modules or []
        self.references = []

    def add_module(self, module):
        assert module not in self.modules
        self.modules.append(module)

    def add_reference(self, guid, major, minor):
        self.references.append((guid, major, minor))

    def export(self, path):
        assert os.path.exists(path) and os.path.isdir(path)
        filenames = []
        for module in self.modules:
            filenames.append(module.export(path))
        return filenames
