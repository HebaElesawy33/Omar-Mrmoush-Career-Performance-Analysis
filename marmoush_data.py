import pandas as pd
import matplotlib.pyplot as plt
import os

# 1. تحديد مكان الملف (على الديسكتوب في مجلد OneDrive)
# تأكدي إن اسم الملف في الإكسيل عندك هو marmoush_data.xlsx
user_profile = os.environ['USERPROFILE']
file_path = os.path.join(user_profile, 'OneDrive', 'Desktop', 'marmoush_data.xlsx')

try:
    # 2. قراءة البيانات من الإكسيل
    df = pd.read_excel(file_path)
    
    # 3. تنظيف البيانات (Data Cleaning)
    # هنشيل الصفوف اللي فيها كلمة Clubs أو Total عشان نقارن الأندية بس
    # وهنشيل أي صفوف مفيهاش أهداف (NaN)
    df_clean = df[~df['Squad'].str.contains('Clubs|Total|Matches', na=False)].copy()
    df_clean = df_clean.dropna(subset=['Squad', 'Gls'])

    # 4. إعداد الرسم البياني (Visualization)
    plt.figure(figsize=(12, 7))
    
    # رسم الأعمدة بلون أخضر رياضي مميز
    bars = plt.bar(df_clean['Squad'], df_clean['Gls'], color='#2ecc71', edgecolor='#27ae60', label='Goals')

    # إضافة الأرقام فوق كل عمود عشان اللي يشوفها يعرف الأهداف بالظبط
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), 
                 ha='center', va='bottom', fontweight='bold', color='#2c3e50')

    # 5. العناوين بالإنجليزية فقط (بدون أي عربي)
    plt.title('Omar Marmoush: Career Goals by Club', fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('Football Club', fontsize=12, fontweight='bold')
    plt.ylabel('Number of Goals', fontsize=12, fontweight='bold')
    
    # تحسين شكل المحور الأفقي (عشان الأسامي متدخلش في بعضها)
    plt.xticks(rotation=30, ha='right')
    
    # إضافة شبكة خفيفة في الخلفية
    plt.grid(axis='y', linestyle='--', alpha=0.3)
    
    # 6. حفظ الصورة النهائية بجودة عالية
    output_path = os.path.join(user_profile, 'OneDrive', 'Desktop', 'Marmoush_Final_Project.png')
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    
    # عرض الرسمة على الشاشة
    plt.show()
    
    print(f"Done! Your professional chart is saved on Desktop as: Marmoush_Final_Project.png")

except Exception as e:
    print(f"An error occurred: {e}")
    print("Hint: Make sure the Excel file is CLOSED before running the code.")