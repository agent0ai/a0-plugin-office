"""Lifecycle-owned registration; dependencies provision automatically on first use."""
def install():
    from helpers import virtual_desktop_routes
    from usr.plugins.office.helpers.viewer_gateway import register
    virtual_desktop_routes.install_route_hooks()
    register()
    return {'sandbox':True, 'installed_on_demand':True}


def disable():
    from usr.plugins.office.helpers.runtime import get_runtime
    get_runtime().stop()


def pre_update():
    disable()


def uninstall():
    disable()
    import shutil
    from usr.plugins.office.helpers.runtime import get_manager, ipc_directory, NAME
    from usr.plugins.office.helpers.viewer_gateway import unregister
    manager = get_manager()
    manager.delete(NAME, missing_ok=True)
    shutil.rmtree(ipc_directory(manager),ignore_errors=True)
    unregister()
