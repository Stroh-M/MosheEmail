class NoInfoError(Exception):
    pass

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