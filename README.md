# SharpExcImpersonation
## 0x01 背景
某一天睡梦中大脑给我提了一个需求: 能不能不依赖Exchange接口模拟任意邮箱用户？
1. 无法访问目标Exchange 接口（OWA、powershell、EWS等）
2. 可访问域控 389 (LDAP) 端口。（除此之外其它任何端口都不能访问）
3. 有高权限账号

### 如何实现：
1. 将用户用户添加进域管组    =>  动作太大且不优雅
2. 将用户添加到Organization Management组  =>  实际上虽然添加到该组并不具ApplicationImpersonation 权限，多次刷新也不行
> 如下是使用exchangelib测试临时添加到`Organization Management `组的用户是否具有`ApplicationImpersonation`权限
![image](https://github.com/user-attachments/assets/07ad3b81-f03b-4458-a74b-0dcb5fcf151c)

4. 复制具有ApplicationImpersonation权限用户LDAP给目标用户

执行 `New-ManagementRoleAssignment -Name "ImpersonationRole" -Role "ApplicationImpersonation" -User "serviceAccount"`并监控LDAP改动信息，使用ldap工具复制已经具有`ApplicationImpersonation`权限用户的相关LDAP信息。这样就能达到添加`ApplicationImpersonation ` 权限的目的。

### 关键点：
** 修改 Active Directory 中的对象属性 **
msExchRoleAssignment 对象：在 Active Directory 中创建一个新的 `msExchRoleAssignment` 对象，该对象链接到要授予` ApplicationImpersonation` 角色的用户。
● msExchRoleLink：此属性指定与角色分配关联的管理角色（即`ApplicationImpersonation` ）。
● msExchUserLink：此属性指定将被授权执行操作的用户或服务帐户的 DN。
● msExchRoleAssignmentFlags：这是一个标志，用于指定角色分配的行为和特性。
● msExchVersion：指定 Exchange 对象的版本信息。

![image](https://github.com/user-attachments/assets/30f6001e-8107-4be2-92f5-287889b84325)

![image](https://github.com/user-attachments/assets/7bd1bed6-743d-42c2-886e-2f25f32dfc44)

![image](https://github.com/user-attachments/assets/0f722804-c47d-48e1-9f2b-040a6e6a2cea)


LDAP 监控工具：https://github.com/p0dalirius/LDAPmonitor
Exchange 中 ApplicationImpersonation 权限详解
ApplicationImpersonation 是 Exchange Server 中的一种角色，授予拥有该角色的用户或服务帐户代表其他用户访问邮箱的能力。这在开发应用程序时非常有用，比如需要访问多个用户的邮箱以进行自动化处理、邮件归档或数据分析等。
## 0x02 代码实现
使用工具操作太过繁琐，也不优雅
注意事项：
1. 确保键值一致
2. 确保数据类型一定要对
`integer8` 对应`IADsLargeInteger` 需要`ActiveDS `库
![image](https://github.com/user-attachments/assets/1060ebc2-4de6-4412-9fbf-8a0891091e12)
![image](https://github.com/user-attachments/assets/46167a2c-db08-4802-9ac9-7704b3acae1c)





