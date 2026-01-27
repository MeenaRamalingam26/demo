
import sqlite3

def delete_all_users():
    try:
        conn = sqlite3.connect('users.db')  # Connect to your DB
        c = conn.cursor()
        
        # Delete all rows from users table
        c.execute("DELETE FROM users")
        
        conn.commit()
        print("✅ All user records have been deleted successfully.")
    except Exception as e:
        print("❌ Error:", e)
    finally:
        conn.close()

# Run the function
delete_all_users()
