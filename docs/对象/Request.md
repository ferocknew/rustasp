## Request 对象

当浏览器向服务器请求页面时，这个行为就被称为一个 request（请求）。Request 对象用于从用户那里获取信息。它的集合、属性和方法描述如下：

### 集合

| 集合                                                         | 描述                                                   |
| :----------------------------------------------------------- | :----------------------------------------------------- |
| ClientCertificate                                            | 包含了存储在客户证书中的所有的字段值。                 |
| [Cookies](https://www.runoob.com/asp/coll-cookies-request.html) | 包含了 HTTP 请求中发送的所有的 cookie 值。             |
| [Form](https://www.runoob.com/asp/coll-form.html)            | 包含了使用 post 方法由表单发送的所有的表单（输入）值。 |
| [QueryString](https://www.runoob.com/asp/coll-querystring.html) | 包含了 HTTP 查询字符串中所有的变量值。                 |
| [ServerVariables](https://www.runoob.com/asp/coll-servervariables.html) | 包含了所有的服务器变量值。                             |

### 属性

| 属性                                                         | 描述                                   |
| :----------------------------------------------------------- | :------------------------------------- |
| [TotalBytes](https://www.runoob.com/asp/prop-totalbytes.html) | 返回在请求正文中客户端发送的字节总数。 |

### 方法

| 方法                                                         | 描述                                                         |
| :----------------------------------------------------------- | :----------------------------------------------------------- |
| [BinaryRead](https://www.runoob.com/asp/met-binaryread.html) | 取回作为 post 请求的一部分而从客户端发送至服务器的数据，并把它存储在一个安全的数组中。 |