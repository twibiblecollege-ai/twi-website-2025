# TWI Bible College Website

**Thy Word International Bible College** — Balanga, Bataan, Philippines  
Built with: HTML · CSS · Vanilla JS | Deployable on **Vercel** | Backend-ready for **Supabase**

---

## 📁 Project Structure

```
twibc-website/
├── index.html               ← Home page
├── about-twibc.html         ← About page
├── academic-programs.html   ← Programs page
├── admissions.html          ← Admissions page
├── registration.html        ← Student registration form
├── campuses.html            ← Global campuses
├── student-life.html        ← Student life & chapel
├── alumni.html              ← Alumni & 25th anniversary
├── contact-us.html          ← Contact page
├── student-portal.html      ← Student portal (Google Script backend)
├── instructor-portal.html   ← Instructor portal (Google Script backend)
├── vercel.json              ← Vercel deployment config
└── README.md
```

---

## 🖥️ VS Code Setup

### 1. Open the project
```bash
code .
```

### 2. Recommended Extensions
Install these from the VS Code Extensions panel (`Ctrl+Shift+X`):

| Extension | Purpose |
|---|---|
| **Live Server** (Ritwick Dey) | Local dev server with hot reload |
| **Prettier** | Code formatting |
| **HTML CSS Support** | Autocomplete |
| **Path Intellisense** | File path hints |

### 3. Launch locally
- Right-click `index.html` → **Open with Live Server**
- Or press `Alt+L Alt+O`
- Opens at `http://127.0.0.1:5500`

### 4. VS Code Settings (already included)
The `.vscode/settings.json` suppresses CSS variable warnings:
```json
{
  "css.validate": false,
  "css.lint.unknownAtRules": "ignore",
  "html.validate.styles": false
}
```

---

## 🚀 Deploy to Vercel

### Option A — Vercel CLI (recommended)
```bash
# Install Vercel CLI
npm install -g vercel

# Login
vercel login

# Deploy from project folder
vercel

# Deploy to production
vercel --prod
```

### Option B — GitHub + Vercel Dashboard
1. Push this project to a GitHub repository:
```bash
git init
git add .
git commit -m "Initial TWIBC website"
git remote add origin https://github.com/YOUR_USERNAME/twibc-website.git
git push -u origin main
```
2. Go to [vercel.com](https://vercel.com) → **New Project**
3. Import your GitHub repository
4. Framework Preset: **Other** (static)
5. Root Directory: `/` (leave default)
6. Click **Deploy** ✅

The `vercel.json` handles all routing automatically.

---

## 🗄️ Supabase Integration (Optional)

The student and instructor portals currently use **Google Apps Script** as a backend. To migrate to Supabase:

### 1. Create a Supabase project
- Go to [supabase.com](https://supabase.com) → New Project
- Note your **Project URL** and **anon public key**

### 2. Create the students table
```sql
create table students (
  id uuid default gen_random_uuid() primary key,
  student_id text unique not null,
  surname text,
  first_name text,
  middle_name text,
  email text unique,
  password text,
  address text,
  mobile text,
  date_of_birth date,
  sex text,
  civil_status text,
  church_name text,
  pastor text,
  classification text,
  subjects_enrolled text[],
  profile_picture_url text,
  created_at timestamp default now()
);
```

### 3. Create the grades table
```sql
create table grades (
  id uuid default gen_random_uuid() primary key,
  student_id text references students(student_id),
  subject_code text,
  subject_name text,
  grade numeric(4,2),
  semester text,
  school_year text,
  instructor_id text,
  created_at timestamp default now()
);
```

### 4. Add Supabase client to portals
Add this snippet to `student-portal.html` and `instructor-portal.html`:
```html
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<script>
  const { createClient } = supabase;
  const supabaseClient = createClient(
    'https://YOUR_PROJECT_ID.supabase.co',
    'YOUR_ANON_PUBLIC_KEY'
  );
</script>
```

### 5. Replace Google Script calls
Replace `google.script.run.checkLogin(...)` calls with:
```javascript
const { data, error } = await supabaseClient
  .from('students')
  .select('*')
  .eq('student_id', studentId)
  .eq('password', password)
  .single();
```

---

## 📞 Contact

**Rev. Orlando N. Acda** — TWI Director  
📞 0920 929 0388  
✉️ twibiblecollege@gmail.com  
📍 Balanga, Bataan, Philippines

---

## 🔗 All Pages

| Page | File | URL |
|---|---|---|
| Home | `index.html` | `/` |
| About | `about-twibc.html` | `/about-twibc.html` |
| Programs | `academic-programs.html` | `/academic-programs.html` |
| Admissions | `admissions.html` | `/admissions.html` |
| Registration | `registration.html` | `/registration.html` |
| Campuses | `campuses.html` | `/campuses.html` |
| Student Life | `student-life.html` | `/student-life.html` |
| Alumni | `alumni.html` | `/alumni.html` |
| Contact | `contact-us.html` | `/contact-us.html` |
| Student Portal | `student-portal.html` | `/student-portal.html` |
| Instructor Portal | `instructor-portal.html` | `/instructor-portal.html` |

---

*To God be all the glory! — TWIBC*
