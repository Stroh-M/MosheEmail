import re

class NoInfoError(Exception):
    # Base class
    @classmethod
    def get_error_name(self):
        return self.__name__
    

class No_Tracking_Number(NoInfoError):
    def __init__(self, message, details=None):
        super().__init__(message)
        self.details = details

    

    def __str__(self):
        return f'Error: {self.args[0]}'

class No_Order_Number(NoInfoError):
    def __init__(self, message, details=None):
        super().__init__(message)
        self.details = details

    def __str__(self):
        return f'Error: {self.args[0]}'  
    
class No_Shipping_Address(NoInfoError):
    def __init__(self, message, details=None):
        super().__init__(message)
        self.details = details

    def __str__(self):
        return f'Error: {self.args[0]}'

# def some_function(value):
#     if value < 0:
#         raise NoOrderNumber("value less then 0")
#     return value

# try:
#     # result = some_function(-5)
#     # print(result)
#     print(re.sub(r'_', ' ', No_Tracking_Number.get_error_name()))
# except NoOrderNumber as non:
#     print(non)
# except No_Tracking_Number as ntn:
#     print()