import requests
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings('ignore')

# 凭证
username = 'schhs334'
password = 'schhs334'

session = requests.Session()
session.verify = False

print('🔍 连接 SMS 系统...')
# 登录
login_url = 'http://sms.chhsban.edu.my/sms/index.php?r=site/login'
login_data = {'LoginForm[username]': username, 'LoginForm[password]': password}
resp = session.post(login_url, data=login_data, timeout=15)

# 获取活动页面
ACTIVITY_PAGE = 'http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create'
resp = session.get(ACTIVITY_PAGE, timeout=15)

if resp.status_code == 200:
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找班级选择下拉框
    class_select = soup.select_one('select#StudentPerformanceM_class_id')
    
    if class_select:
        print('\n✅ 找到班级选择下拉框！\n')
        options = class_select.select('option[value]')
        
        print(f'系统中共有 {len(options)} 个班级:\n')
        
        # 找 J1B, S3A, S3B
        target_classes = ['J1B', 'S3A', 'S3B']
        found = []
        
        for option in options:
            class_name = option.get_text(strip=True)
            class_id = option.get('value', '')
            
            for target in target_classes:
                if target in class_name:
                    print(f'✓ {class_name:20} -> ID: {class_id}')
                    found.append((class_name, class_id))
        
        if found:
            print(f'\n📌 找到的缺失班级：')
            for name, id in found:
                print(f'  {name}: {id}')
        else:
            print(f'\n⚠ 未找到 {target_classes} 这些班级')
            print('\n所有班级列表:')
            for option in options[:30]:  # 显示前 30 个
                class_name = option.get_text(strip=True)
                class_id = option.get('value', '')
                print(f'  {class_name:20} -> {class_id}')
    else:
        print('❌ 无法找到班级选择下拉框')
else:
    print(f'❌ 连接失败: {resp.status_code}')
