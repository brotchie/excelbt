from win32com.client import Dispatch

def pytest_funcarg__xl(request):
    xl = Dispatch('Excel.Application')

    def finalize():
        xl.DisplayAlerts = 0
        xl.Quit()

    request.addfinalizer(finalize)
    return xl
