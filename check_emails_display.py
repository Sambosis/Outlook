"""
Diagnostic script to check why emails aren't displaying
"""
import os
import sys

print("=" * 60)
print("EMAIL DISPLAY DIAGNOSTICS")
print("=" * 60)

# Check 1: Database configuration
print("\n1. Checking database configuration...")
try:
    from db import get_database_url, SessionLocal
    db_url = get_database_url()
    print(f"   ✓ Database URL: {db_url[:50]}...")
    
    # Check if it's SQLite
    if db_url.startswith('sqlite'):
        db_file = db_url.replace('sqlite:///', '')
        if os.path.exists(db_file):
            print(f"   ✓ SQLite database exists: {db_file}")
            print(f"   ✓ File size: {os.path.getsize(db_file):,} bytes")
        else:
            print(f"   ✗ SQLite database NOT found: {db_file}")
            print(f"   → You need to initialize the database or fetch emails")
    else:
        print(f"   ✓ Using PostgreSQL (production)")
        
except Exception as e:
    print(f"   ✗ Error checking database: {e}")
    sys.exit(1)

# Check 2: Database tables
print("\n2. Checking database tables...")
try:
    from models import Email, Attachment, Base
    from db import engine
    
    # Check if tables exist
    inspector = None
    try:
        from sqlalchemy import inspect
        inspector = inspect(engine)
        tables = inspector.get_table_names()
        print(f"   ✓ Tables in database: {tables}")
        
        if 'emails' not in tables:
            print(f"   ✗ 'emails' table not found!")
            print(f"   → Run: python -c \"from app import init_db; init_db()\"")
    except Exception as e:
        print(f"   ! Could not inspect tables: {e}")
        
except Exception as e:
    print(f"   ✗ Error checking tables: {e}")

# Check 3: Email count
print("\n3. Checking email count...")
try:
    from db import SessionLocal
    from models import Email
    
    with SessionLocal() as session:
        count = session.query(Email).count()
        print(f"   ✓ Total emails in database: {count}")
        
        if count == 0:
            print(f"   ! No emails found in database")
            print(f"   → Options:")
            print(f"      1. Use 'Check for New Emails' button in the UI")
            print(f"      2. POST to /check-emails endpoint")
            print(f"      3. Manually import emails")
        else:
            # Show sample
            sample = session.query(Email).order_by(Email.datetime_received.desc()).first()
            if sample:
                print(f"   ✓ Most recent email:")
                print(f"      Subject: {sample.subject}")
                print(f"      From: {sample.sender}")
                print(f"      Date: {sample.datetime_received}")
                
except Exception as e:
    print(f"   ✗ Error querying emails: {e}")
    import traceback
    traceback.print_exc()

# Check 4: Environment variables
print("\n4. Checking Exchange configuration...")
try:
    from dotenv import load_dotenv
    load_dotenv()
    
    required_vars = [
        'EXCHANGE_EMAIL',
        'EXCHANGE_DOMAIN_USERNAME', 
        'EXCHANGE_PASSWORD',
        'EXCHANGE_SERVER',
        'EXCHANGE_VERSION',
        'TIMEZONE',
        'DAYS_AGO'
    ]
    
    missing = []
    for var in required_vars:
        if os.getenv(var):
            print(f"   ✓ {var} is set")
        else:
            print(f"   ✗ {var} is NOT set")
            missing.append(var)
    
    if missing:
        print(f"\n   ✗ Missing variables: {missing}")
        print(f"   → Cannot fetch emails from Exchange server")
    else:
        print(f"\n   ✓ All Exchange variables configured")
        
except Exception as e:
    print(f"   ✗ Error checking environment: {e}")

# Check 5: Template file
print("\n5. Checking template files...")
try:
    template_path = "templates/index.html"
    if os.path.exists(template_path):
        print(f"   ✓ Template exists: {template_path}")
        
        # Check if it has the emails loop
        with open(template_path, 'r', encoding='utf-8') as f:
            content = f.read()
            if '{% for email in emails %}' in content:
                print(f"   ✓ Template has email loop")
            else:
                print(f"   ✗ Template missing email loop")
                
            if '{% if emails %}' in content:
                print(f"   ✓ Template checks for emails")
            else:
                print(f"   ! Template doesn't check if emails exist")
    else:
        print(f"   ✗ Template NOT found: {template_path}")
        
except Exception as e:
    print(f"   ✗ Error checking template: {e}")

# Summary
print("\n" + "=" * 60)
print("NEXT STEPS:")
print("=" * 60)
print("""
If database is empty:
  1. Make sure Exchange credentials are in .env file
  2. Run the app: python app.py
  3. Open http://localhost:5000 in browser
  4. Click 'Check for New Emails' button
  
If database has emails but they don't display:
  1. Check browser console for JavaScript errors
  2. Check that Flask app is serving index.html correctly
  3. Verify the template is rendering the emails variable
  
For Heroku deployment:
  1. Make sure DATABASE_URL is set (PostgreSQL)
  2. Push latest code: git push heroku main:master
  3. Check logs: heroku logs --tail --app your-app-name
  4. Initialize DB: heroku run python -c "from app import init_db; init_db()"
""")
