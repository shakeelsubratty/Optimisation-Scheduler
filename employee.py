class Employee:

    def __init__(self, name, team, role, technicality, max_hours, days_available_weekly):
        self.max_hours = max_hours
        self.name = name
        self.role = role
        self.techincality = technicality
        self.team = team
        self.days_available_weekly = days_available_weekly
        self.allocated_payrolls = []
        self.total_allocated_time = 0

    def get_name(self):
        return self.name

    def get_technicality(self):
        return self.techincality

    def get_max_hours(self):
        return self.max_hours

    def get_days_available_weekly(self):
        return self.days_available_weekly

    def get_allocated_payrolls(self):
        return self.allocated_payrolls

    def get_allocated_payrolls_total_time(self):
        return self.total_allocated_time

    def get_allocated_payrolls_time_7days_from_due_date(self, payroll_due_date):
        total_time = 0
        for payroll in self.allocated_payrolls:
            if payroll.get_due_date() >= payroll_due_date:
                total_time += payroll.get_processing_time()
        return total_time

    def allocate_payrolls(self, payrolls):
        for payroll in payrolls:
            self.allocated_payrolls.append(payroll)
            self.total_allocated_time += payroll.get_processing_time()
