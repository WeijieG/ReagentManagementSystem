# venv\Scripts\activate
import json
import os
import sys
import hashlib
from flask import Flask, send_file, request, jsonify
import ssl  # 导入ssl模块

app = Flask(__name__)
# venv\Scripts\activate

# 配置
UPDATES_DIR = 'updates'
CURRENT_VERSION = '2.1.0'

# 更新说明。放在 updates\changelog.txt 里就不用改代码；
# 没有这个文件时退回到一句通用说明。
CHANGELOG_FILE = os.path.join(UPDATES_DIR, 'changelog.txt')


def load_changelog():
    try:
        with open(CHANGELOG_FILE, 'r', encoding='utf-8') as f:
            text = f.read().strip()
        if text:
            return text
    except OSError:
        pass
    return f"版本 {CURRENT_VERSION} 更新内容：\n- 优化性能\n- 修复已知问题"

# 语义化版本比较。**不要**图省事写成 client_version > CURRENT_VERSION ——
# 那是字符串比较，'2.0.9' > '2.0.10' 为真，会直接回 "Already up to date"，
# 更新永远推不出去。客户端的 app/core/version.py 已经修过同样的问题。
# 版本号一改记得同步 config.ini 的 [Application] version。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app.core.version import compare_versions  # noqa: E402

# 证书文件路径（将在首次运行时自动生成）
CERT_FILE = 'self_signed.crt'
KEY_FILE = 'private.key'

def generate_self_signed_cert():
    """生成自签名证书（如果不存在）"""
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes
    
    # 检查证书是否已存在
    if os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE):
        return
    
    # 生成私钥
    private_key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
        backend=default_backend()
    )
    
    # 生成自签名证书
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, u"localhost"),
    ])
    
    cert = x509.CertificateBuilder().subject_name(
        subject
    ).issuer_name(
        issuer
    ).public_key(
        private_key.public_key()
    ).serial_number(
        x509.random_serial_number()
    ).not_valid_before(
        datetime.datetime.utcnow()
    ).not_valid_after(
        datetime.datetime.utcnow() + datetime.timedelta(days=365)
    ).add_extension(
        x509.SubjectAlternativeName([x509.DNSName(u"localhost")]),
        critical=False,
    ).sign(private_key, hashes.SHA256(), default_backend())
    
    # 保存私钥
    with open(KEY_FILE, "wb") as f:
        f.write(private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))
    
    # 保存证书
    with open(CERT_FILE, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

def calculate_md5(file_path):
    """计算文件的MD5值"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

@app.route('/update.json')
def update_info():
    """提供更新信息"""
    client_version = request.args.get('version', '1.0.1')

    # 检查是否有更新（语义化比较：客户端不低于服务端就没有更新）
    if compare_versions(client_version, CURRENT_VERSION) >= 0:
        return jsonify({
            "update": False,
            "message": "Already up to date"
        })
    
    # 获取最新安装包
    latest_setup = None
    for file in os.listdir(UPDATES_DIR):
        if file.endswith('.exe'):
            latest_setup = file
            break
    
    if not latest_setup:
        return jsonify({"error": "No update package found"}), 404
    
    # 计算MD5
    setup_path = os.path.join(UPDATES_DIR, latest_setup)
    md5 = calculate_md5(setup_path)
    
    # 返回更新信息
    return jsonify({
        "update": True,
        "version": CURRENT_VERSION,
        "url": f"{request.host_url}download/{latest_setup}",
        "md5": md5,
        "changelog": load_changelog(),
        "size": os.path.getsize(setup_path)
    })

@app.route('/download/<filename>')
def download_file(filename):
    """提供安装包下载"""
    file_path = os.path.join(UPDATES_DIR, filename)
    if not os.path.exists(file_path):
        return "File not found", 404
    
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    # 确保更新目录存在
    os.makedirs(UPDATES_DIR, exist_ok=True)
    
    # 生成自签名证书（如果不存在）
    try:
        import datetime  # 延迟导入避免不需要时依赖
        from cryptography.hazmat.primitives import serialization
        generate_self_signed_cert()
        ssl_context = (CERT_FILE, KEY_FILE)
    except ImportError:
        print("缺少加密依赖，请安装：pip install cryptography")
        ssl_context = 'adhoc'  # 使用Flask内置的临时证书
    
    # 启动HTTPS服务器
    app.run(
        host='0.0.0.0', 
        port=8000, 
        threaded=True,
        ssl_context=ssl_context  # 启用HTTPS
    )