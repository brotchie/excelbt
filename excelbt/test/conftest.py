def pytest_funcarg__xl(request):
    from win32com.client import Dispatch
    xl = Dispatch('Excel.Application')

    def finalize():
        xl.DisplayAlerts = 0
        xl.Quit()

    request.addfinalizer(finalize)
    return xl
