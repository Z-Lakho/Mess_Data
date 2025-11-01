def get_mess_charges(group, designation):
    group = str(group)
    designation = str(designation)
    if 'Trainee' in group or 'Management Trainee' in designation:
        return 5000
    elif 'Worker' in group:
        return 2800
    elif 'Officer' in group or 'Consultant Officer' in group:
        if 'Deputy General Manager' in designation:
            return 6000
        elif 'Lt Col' in designation:
            return 4500
        elif 'Managing Director' in designation:
            return 3000
        elif 'Admin Coordinator' in designation:
            return 4200
        else:
            return 3000
    return 3000  # Default charges