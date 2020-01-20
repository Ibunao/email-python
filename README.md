## 配置

```python
# 配置发件人和发件邮箱 例 {'name':'张三', 'email':'zhangsan@163.com'}
self.from_addr = {'name':'发件人', 'email':'邮箱'}
# 邮箱授权码,注意这里不是邮箱密码
self.password = '***'
# SMTP服务器，这里默认为 163 的
self.smtp_server = 'smtp.163.com'
```



## 使用

可在其他类引入后使用

最下面有使用的 Dome 参考，