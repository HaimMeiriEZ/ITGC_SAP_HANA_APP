def check_status(user_role: str) -> bool:
    """פונקציית בדיקה לבדיקת אוטומציית ClickUp v2"""
    if user_role == "admin":
        return True
    return False

# בדיקה נוספת
print("ClickUp v2 test active")
