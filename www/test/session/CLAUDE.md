# Session 测试计划

## 测试目标

验证 ASP-lite 的 Session 对象功能是否符合 Classic ASP 的 Session 规范。

## Session 对象功能清单

### 基础功能
- [ ] `Session.Contents(key)` - 获取/设置 Session 值
- [ ] `Session(key)` - 简写形式
- [ ] `Session.SessionID` - 获取唯一会话标识符
- [ ] `Session.Timeout` - 获取/设置会话超时时间（分钟）

### 集合功能
- [ ] `Session.Contents.Count` - 获取 Session 变量数量
- [ ] `Session.Contents.Item(key)` - 获取指定键的值
- [ ] `Session.Contents.Remove(key)` - 删除指定键
- [ ] `Session.Contents.RemoveAll()` - 清空所有 Session 变量
- [ ] `Session.Contents.Key(index)` - 通过索引获取键名

### 方法功能
- [ ] `Session.Abandon()` - 立即销毁当前会话
- [ ] `Session.IsClientConnected` - 检查客户端是否连接

### 事件功能（可能不支持）
- [ ] `Session_OnStart` - 会话开始事件
- [ ] `Session_OnEnd` - 会话结束事件

## 测试用例设计

### 基础存储测试
1. **test_001_session_basic.asp**
   - 设置 Session 变量
   - 读取 Session 变量
   - 修改 Session 变量
   - 验证值在请求之间保持

2. **test_002_session_types.asp**
   - 测试不同数据类型：字符串、数字、布尔值、日期
   - 验证类型正确保存和恢复

3. **test_003_session_objects.asp**
   - 测试存储简单对象（如数组）
   - 验证对象序列化

### SessionID 测试
4. **test_004_session_id.asp**
   - 获取 Session.SessionID
   - 验证 ID 在会话期间不变
   - 验证不同会话 ID 不同

### 超时测试
5. **test_005_session_timeout.asp**
   - 设置 Session.Timeout
   - 读取 Session.Timeout
   - 验证超时后 Session 清空

### 集合操作测试
6. **test_006_session_contents.asp**
   - Session.Contents.Count
   - Session.Contents.Key(index)
   - 遍历所有 Session 变量

7. **test_007_session_remove.asp**
   - Session.Contents.Remove(key)
   - Session.Contents.RemoveAll()
   - 验证删除后变量不存在

### Abandon 测试
8. **test_008_session_abandon.asp**
   - 调用 Session.Abandon()
   - 验证会话被销毁
   - 验证新请求获得新会话

### 跨页面测试
9. **test_009_session_cross_page.asp**
   - 页面 A 设置 Session
   - 页面 B 读取 Session
   - 验证跨页面共享

### 存储模式测试
10. **test_010_session_persistence.asp**
    - 测试 memory 模式
    - 测试 json 模式
    - 测试 redis 模式（如果可用）
    - 验证重启后数据持久化（json/redis）

## 测试环境配置

### .env 配置
```env
# Session 存储模式：memory/json/redis
SESSION_STORAGE_MODE=memory

# JSON 文件路径（json 模式）
SESSION_JSON_FILE=./data/sessions.json

# Redis 配置（redis 模式）
REDIS_URL=redis://127.0.0.1:6379

# Session 超时时间（分钟）
SESSION_TIMEOUT=20
```

## 测试执行步骤

### 1. 基础功能测试
```bash
# 启动服务器
cargo run

# 访问测试页面
open http://localhost:8080/test/session/test_001_session_basic.asp
```

### 2. 跨页面测试
1. 打开页面 A，设置 Session
2. 记录 SessionID
3. 打开页面 B，读取 Session
4. 验证 SessionID 一致
5. 验证 Session 值正确

### 3. 超时测试
1. 设置 Session.Timeout = 1（1分钟）
2. 设置 Session 变量
3. 等待 1 分钟
4. 刷新页面，验证 Session 清空

### 4. 持久化测试
1. 切换到 json 模式
2. 设置 Session 变量
3. 重启服务器
4. 刷新页面，验证 Session 保持

## 预期结果

### Memory 模式
- ✅ Session 存储在内存中
- ❌ 重启后数据丢失
- ✅ 性能最好

### JSON 模式
- ✅ Session 存储在 JSON 文件
- ✅ 重启后数据保持
- ⚠️ 性能中等
- ⚠️ 并发安全性待验证

### Redis 模式
- ✅ Session 存储在 Redis
- ✅ 重启后数据保持
- ✅ 支持多服务器共享
- ✅ 性能好
- ⚠️ 需要 Redis 服务

## 已知限制

1. **不支持复杂对象**：Session 只支持简单数据类型
2. **不支持 Session 事件**：OnStart/OnEnd 可能不支持
3. **数组支持有限**：可能只支持一维数组
4. **对象不支持**：不支持存储 COM/ActiveX 对象

## 测试报告模板

```
## 测试执行报告

**日期**: YYYY-MM-DD
**测试人员**: xxx
**Session 模式**: memory/json/redis

### 测试结果

| 用例编号 | 用例名称 | 状态 | 备注 |
|---------|---------|------|------|
| test_001 | Session 基础存储 | ✅/❌ | |
| test_002 | Session 数据类型 | ✅/❌ | |
| ... | ... | ... | |

### 发现的问题

1. 问题描述
   - 复现步骤
   - 预期结果
   - 实际结果
   - 错误信息

### 改进建议

...
```
