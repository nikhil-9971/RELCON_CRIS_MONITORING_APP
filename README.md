# RO Code Report Merger — Desktop App

Duplicate RO Code rows ko ek clean row me merge karta hai — Engineer ke liye ready report.

## Kya merge hota hai
- **Same rehte hain** (pehla value liya jata hai): Zone, Region, Salesarea, Vendor, RO Code, RO Sap Code, RO Name, **ROC Date**
- **Comma se combine hote hain** (unique values): Ticket No, Ageing Days, Device, Device Details, Current Dependency, Automation Vendor Comments, HPCL Comments
- Ek extra **"Ticket Count"** column bhi add hota hai — kitne tickets us RO Code ke merge hue

## App run karna (aapke apne computer par, testing ke liye)
```bash
pip install -r requirements.txt
python desktop_app.py
```

## Windows .exe banana

### Option A — GitHub Actions se (Windows machine ki zaroorat nahi, FREE)
Is folder me `.github/workflows/build-exe.yml` already diya hua hai jo cloud me
automatically Windows .exe bana deta hai.

1. GitHub par ek naya (private ya public) repository banao
2. Is poore folder (`desktop_app.py`, `requirements.txt`, `.github` folder) ko us repo me push karo:
   ```bash
   git init
   git add .
   git commit -m "RO Code Merger app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
3. GitHub repo ke **"Actions"** tab me jao — build automatically start ho jayega (2-3 min lagte hain)
4. Build complete hone ke baad, us workflow run ke andar **"Artifacts"** section me
   `RO_Code_Merger-windows-exe` milega — usko download karo, andar `.exe` hoga
5. Ye `.exe` kisi bhi Windows user ko bhej sakte ho — unke system par Python install
   hone ki zaroorat nahi

### Option B — Apni Windows machine par (agar available ho)
1. Windows machine par requirements install karo: `pip install -r requirements.txt`
2. Ye command chalao:
   ```bash
   pyinstaller --noconsole --onefile --name "RO_Code_Merger" desktop_app.py
   ```
3. `dist` folder ke andar `RO_Code_Merger.exe` ban jayega — yahi file kisi ko bhi bhej sakte ho.

> **Note:** Mera sandbox Linux hai, isliye main khud yahan se seedha Windows `.exe`
> generate nahi kar sakta (PyInstaller cross-platform build support nahi karta).
> Isliye Option A (GitHub Actions) sabse aasan tarika hai — bina apni Windows
> machine ke, cloud me hi asli `.exe` ban jayega.

## Note
- `.exe` build hamesha usi OS par karo jis OS ke liye use chahiye (Windows machine → Windows .exe, Mac → Mac app).
- Icon add karna ho to: `pyinstaller --noconsole --onefile --icon=your_icon.ico --name "RO_Code_Merger" desktop_app.py`
