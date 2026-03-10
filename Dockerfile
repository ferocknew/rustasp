# 使用最小化的 Alpine Linux 作为基础镜像
FROM alpine:3.19

# 安装必要的运行时依赖
RUN apk add --no-cache ca-certificates tzdata

# 创建非 root 用户
RUN addgroup -g 1000 vbscript && \
    adduser -u 1000 -G vbscript -s /bin/sh -D vbscript

# 设置工作目录
WORKDIR /app

# 复制编译好的二进制文件
COPY vbscript /app/vbscript

# 复制默认配置和示例文件
COPY .env.example /app/.env.example
COPY www /app/www
COPY README.md /app/README.md
COPY VERSION /app/VERSION

# 创建运行时目录
RUN mkdir -p /app/runtime/sessions && \
    chown -R vbscript:vbscript /app

# 切换到非 root 用户
USER vbscript

# 暴露端口
EXPOSE 8080

# 设置默认环境变量
ENV HOME_DIR=/app/www
ENV PORT=8080

# 启动服务
CMD ["./vbscript"]
