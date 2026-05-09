from app.session import get_current_user


ROLE_PERMISSIONS = {
    "admin": {
        "dashboard",
        "stock_view",
        "movement_entry",
        "batch_manager",
        "fefo_helper",
        "audit_viewer",
        "user_manager",
        "purchase_orders",
        "supplier_deliveries",
        "backup_restore",
        "reports",
        "forecast",
    },

    "pharmacy_manager": {
        "dashboard",
        "stock_view",
        "movement_entry",
        "batch_manager",
        "fefo_helper",
        "audit_viewer",
        "purchase_orders",
        "supplier_deliveries",
        "reports",
        "forecast",
    },

    "pharmacist": {
        "dashboard",
        "stock_view",
        "movement_entry",
        "batch_manager",
        "fefo_helper",
        "forecast",
    },

    "storekeeper": {
        "dashboard",
        "stock_view",
        "movement_entry",
        "batch_manager",
        "supplier_deliveries",
    },

    "procurement": {
        "dashboard",
        "stock_view",
        "purchase_orders",
        "supplier_deliveries",
        "reports",
        "forecast",
    },

    "viewer": {
        "dashboard",
        "stock_view",
        "reports",
        "forecast",
    },
}


def current_role():
    user = get_current_user()
    return user.get("role", "viewer")


def user_can(permission_name):
    role = current_role()
    allowed = ROLE_PERMISSIONS.get(role, set())
    return permission_name in allowed


def get_denied_message(permission_name):
    role = current_role()
    return (
        f"Access denied.\n\n"
        f"Your role: {role}\n"
        f"Required permission: {permission_name}\n\n"
        f"هذا المستخدم لا يملك صلاحية تنفيذ هذه العملية."
    )


if __name__ == "__main__":
    print("Current role:", current_role())
    print("Permissions:")
    for permission in sorted(ROLE_PERMISSIONS.get(current_role(), set())):
        print("-", permission)
