class Employee:

    def __init__(self, name, team, role, technicality, max_hours, days_available_weekly):
        self.max_hours = max_hours
        self.name = name
        self.role = role
        self.techincality = technicality
        self.team = team
        self.days_available_weekly = days_available_weekly

    def get_name(self):
        return self.name

    def get_technicality(self):
        return self.techincality

    def get_max_hours(self):
        return self.max_hours

    def get_days_available_weekly(self):
        return self.days_available_weekly
