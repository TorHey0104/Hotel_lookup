ROLE_KEYS = ["AVP", "MD", "GM", "Engineering", "DOF", "RegionalEngineeringSpecialist"]

def get_role_map(roles_cfg, defaults):
    """Return mapping of role -> delivery mode."""
    role_send = {}
    for role in ROLE_KEYS:
        role_send[role] = roles_cfg.get(role, defaults.get(role, "Skip"))
    return role_send
