from helpers.extension import Extension
from usr.plugins.office import hooks

class OfficeStartup(Extension):
    def execute(self, **kwargs):
        hooks.install()
