def check_status(user_role: str) -> bool:
    """פונקציית בדיקה לבדיקת אוטומציית ClickUp"""
    if user_role == "admin":
        return True
    return False
