## Response 对象

ASP Response 对象用于从服务器向用户发送输出的结果。它的集合、属性和方法描述如下：

### 集合

| 集合                                                         | 描述                                                         |
| :----------------------------------------------------------- | :----------------------------------------------------------- |
| [Cookies](https://www.runoob.com/asp/coll-cookies-response.html) | 设置 cookie 的值。如果 cookie 不存在，则创建 cookie ，并设置指定的值。 |

### 属性

| 属性                                                         | 描述                                                       |
| :----------------------------------------------------------- | :--------------------------------------------------------- |
| [Buffer](https://www.runoob.com/asp/prop-buffer.html)        | 规定是否缓冲页面的输出。                                   |
| [CacheControl](https://www.runoob.com/asp/prop-cachecontrol.html) | 设置代理服务器是否可以缓存由 ASP 产生的输出。              |
| [Charset](https://www.runoob.com/asp/prop-charset.html)      | 将字符集的名称追加到 Response 对象中的 content-type 报头。 |
| [ContentType](https://www.runoob.com/asp/prop-contenttype.html) | 设置 Response 对象的 HTTP 内容类型。                       |
| [Expires](https://www.runoob.com/asp/prop-expires.html)      | 设置页面在失效前的浏览器缓存时间（分钟）。                 |
| [ExpiresAbsolute](https://www.runoob.com/asp/prop-expiresabsolute.html) | 设置浏览器上页面缓存失效的日期和时间。                     |
| [IsClientConnected](https://www.runoob.com/asp/prop-isclientconnected.html) | 指示客户端是否已从服务器断开。                             |
| [Pics](https://www.runoob.com/asp/prop-pics.html)            | 向 response 报头的 PICS 标签追加值。                       |
| [Status](https://www.runoob.com/asp/prop-status.html)        | 规定由服务器返回的状态行的值。                             |

### 方法

| 方法                                                         | 描述                                         |
| :----------------------------------------------------------- | :------------------------------------------- |
| [AddHeader](https://www.runoob.com/asp/met-addheader.html)   | 向 HTTP 响应添加新的 HTTP 报头和值。         |
| [AppendToLog](https://www.runoob.com/asp/met-appendtolog.html) | 向服务器日志条目的末端添加字符串。           |
| [BinaryWrite](https://www.runoob.com/asp/met-binarywrite.html) | 在没有任何字符转换的情况下直接向输出写数据。 |
| [Clear](https://www.runoob.com/asp/met-clear.html)           | 清除已缓冲的 HTML 输出。                     |
| [End](https://www.runoob.com/asp/met-end.html)               | 停止处理脚本，并返回当前的结果。             |
| [Flush](https://www.runoob.com/asp/met-flush.html)           | 立即发送已缓冲的 HTML 输出。                 |
| [Redirect](https://www.runoob.com/asp/met-redirect.html)     | 把用户重定向到一个不同的 URL。               |
| [Write](https://www.runoob.com/asp/met-write-response.html)  | 向输出写指定的字符串。                       |