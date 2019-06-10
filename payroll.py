class Payroll:

    def __init__(self, id,  previous_employee, technicality, due_date, pay_date, data_sent_date, processing_time, do_not_reallocate, imported=False):
        self.previous_employee = previous_employee
        self.id = id
        self.pay_date = pay_date
        self.due_date = due_date
        self.data_sent_date = data_sent_date
        self.techincality = technicality
        self.processing_time = processing_time
        self.do_not_reallocate = do_not_reallocate

    def get_id(self):
        return self.id

    def get_technicality(self):
        return self.techincality

    def get_due_date(self):
        return self.due_date

    def get_pay_date(self):
        return self.pay_date

    def get_data_sent_date(self):
        return self.data_sent_date

    def get_processing_time(self):
        return self.processing_time

    def get_do_not_reallocate(self):
        return self.do_not_reallocate

    def get_previous_employee(self):
        return self.previous_employee