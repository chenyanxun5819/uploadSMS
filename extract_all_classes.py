import requests
from bs4 import BeautifulSoup
import warnings
import json
from pathlib import Path

warnings.filterwarnings('ignore')

# 凭证
username = 'schhs334'
password = 'schhs334'

session = requests.Session()
session.verify = False

print('🔍 连接 SMS 系统，提取所有班级...')

# 登录
login_url = 'http://sms.chhsban.edu.my/sms/index.php?r=site/login'
login_data = {'LoginForm[username]': username, 'LoginForm[password]': password}
try:
    resp = session.post(login_url, data=login_data, timeout=15)
    print('✓ 登录成功')
except Exception as e:
    print(f'✗ 登录失败: {e}')
    exit(1)

# 获取活动页面
ACTIVITY_PAGE = 'http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create'
try:
    resp = session.get(ACTIVITY_PAGE, timeout=15)
    print(f'✓ 获取活动页面成功 (状态码: {resp.status_code})')
except Exception as e:
    print(f'✗ 获取页面失败: {e}')
    exit(1)

soup = BeautifulSoup(resp.text, 'html.parser')

# 查找班级选择下拉框
class_select = soup.select_one('select#StudentPerformanceM_class_id')

if not class_select:
    print('❌ 无法找到班级选择下拉框')
    # 尝试其他选择器
    print('尝试其他选择器...')
    class_select = soup.select_one('select[name*="class_id"]')
    if class_select:
        print(f'✓ 找到备用班级选择框: {class_select.get("id", class_select.get("name"))}')
    else:
        print('❌ 找不到任何班级选择框')
        exit(1)

options = class_select.select('option[value]')
print(f'\n✓ 找到 {len(options)} 个班级\n')

class_mapping = {}
for option in options:
    class_name = option.get_text(strip=True)
    class_id = option.get('value', '')
    
    if class_id and class_name:  # 排除空值
        class_mapping[class_name] = class_id
        print(f'  {class_name:20} -> {class_id}')

print(f'\n✓ 共提取 {len(class_mapping)} 个班级')

# 保存到配置目录
config_dir = Path.home() / '.sms_app'
config_dir.mkdir(exist_ok=True)

config_file = config_dir / 'config.json'

# 读取现有配置（如果存在）
if config_file.exists():
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
else:
    config = {}

# 更新班级映射
config['class_mapping'] = class_mapping
config['last_updated'] = str(__import__('datetime').datetime.now())

# 保存
with open(config_file, 'w', encoding='utf-8') as f:
    json.dump(config, f, indent=2, ensure_ascii=False)

print(f'\n✅ 已保存到: {config_file}')
print(f'✅ 共 {len(class_mapping)} 个班级')
