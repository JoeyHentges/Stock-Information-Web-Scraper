import pip
import imp

def main():
    packageList = [] #array to hold package name
    package_list = open('Packages.txt', 'r')
    for line in package_list:
        packageList.append(line.replace('\n', ''))
    package_list.close()
    for i in range(len(packageList)):
        try:
            imp.find_module(packageList[i])
            found = True
        except ImportError:
            found = False
        if found == False:
            install(packageList[i])

def install(package):
    import importlib
    try:
        importlib.import_module(package)
    except ImportError:
        import pip
        pip.main(['install', package])
    finally:
        globals()[package] = importlib.import_module(package)

main()
