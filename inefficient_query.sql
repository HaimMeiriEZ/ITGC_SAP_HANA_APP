-- שאילתה לא אופטימלית לשליפת עסקאות גדולות
SELECT * FROM transactions WHERE user_id = 105 AND amount > 500;
