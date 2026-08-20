# 配置核查 · 麒麟靶机 Docker 镜像

本地起一台「银河麒麟 V10」靶机，用于自测 `kylin/check_kylin.sh` 的核查效果，无需真实麒麟服务器。

## 前置条件

- Docker 20.10+（含 Docker Desktop on Windows/WSL2）
- 构建上下文必须是**项目根目录**（因为要 `COPY kylin/check_kylin.sh`）

## 为什么用 systemctl 替身而不是真 systemd

Kylin 的 systemd 是 **v243、编译参数 `default-hierarchy=legacy`（cgroup v1）**，而 Docker Desktop 的 WSL2 后端是 **cgroup v2**，两者不兼容——容器启动 `/sbin/init` 会秒退 255、零日志。

因此本镜像用 [docker-systemctl-replacement](https://github.com/gdraheim/docker-systemctl-replacement) 的 `systemctl3.py`（已下载到 `docker/systemctl3.py`）替换 `/usr/bin/systemctl`，在**没有 systemd** 的环境下也能执行 `systemctl is-active / is-enabled / start / stop` 等命令。容器以该脚本作为 PID 1 时，会自动进入 init 模式并拉起 enabled 服务。

> 若你在真 Linux 主机（非 Docker Desktop）上跑，想用真 systemd，可把 `CMD` 改回 `["/sbin/init"]` 并加 `-v /sys/fs/cgroup:/sys/fs/cgroup:ro`。

## 构建

```bash
cd 配置核查
docker build -f docker/Dockerfile -t kylin-target:latest .
```

## 运行

```bash
docker run -d --privileged --name kylin-target kylin-target:latest
```

> `--privileged` 用于让 firewalld/iptables/auditd 等检查尽可能真实；即使不加容器也能起，只是这些项会判 `fail`/`manual`。

## 自测核查脚本

```bash
docker exec -it kylin-target bash -c "cd /opt && bash check_kylin.sh"
```

跑完报告生成在容器内 `/opt/output/`，用 `docker cp` 拷出来：

```bash
docker cp kylin-target:/opt/output ./docker/output
```

## 内置组件

| 类别 | 组件 | 容器内实测状态 |
|------|------|--------------|
| 中间件 | Redis、Apache(httpd)、Tomcat | ✅ active（端口 6379/80/8080） |
| 文件共享 | Samba、vsftpd | ✅ active（端口 139/445/21） |
| 数据库 | MariaDB、PostgreSQL | ⚠️ 见下方「已知限制」 |
| 中间件 | Nginx | ⚠️ 与 httpd 抢 80 端口 |
| 系统 | SSH(sshd)、firewalld、auditd、SELinux | ⚠️ 容器环境受限 |
| 核查脚本 | `/opt/check_kylin.sh` | ✅ 手动执行 |

> 达梦（DM8）、人大金仓（KingbaseES）为商业数据库，需授权文件，未内置。请用现成镜像单独起容器，例如 `xuxuclassmate/dameng`、`wephoon/kingbase`。

## 已知限制

1. **MariaDB / PostgreSQL / sshd 起不来**——systemctl 替身对 MariaDB 的 `ExecStartPre=unset-environment`、PostgreSQL 的初始化检查、sshd 的 host key 生成支持不完整，这三项会判 `fail`。如需覆盖这三项，建议改在真 Linux 主机用真 systemd（见上）。
2. **Nginx 与 Apache(httpd) 冲突**——两者默认都绑 80 端口，httpd 先起导致 nginx 报 `bind() ... Address already in use`。只保留了 httpd。
3. **firewalld / auditd / SELinux 在容器内不完整工作**——内核隔离限制，相关项判 `fail`/`manual` 属预期。
4. **网络设备（第 5 章）23 项**本就是网络设备层面，脚本标「不适用」。
5. **Windows XP/7 无法 docker 化**——`win_xp7/check_xp7.vbs` 需在真机或 VirtualBox/VMware 虚拟机里测。

## 若基础镜像换成 apt 系（优麒麟 / Ubuntu-Kylin）

本 Dockerfile 默认按 `macrosan/kylin`（rpm/yum 系）编写。若改用 apt 系镜像（优麒麟等），把 `RUN` 里的 yum 段替换为：

```bash
RUN set -eux; \
    apt-get update; \
    DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
        systemd systemd-sysv sudo passwd iproute2 net-tools hostname procps \
        cron rsyslog logrotate util-linux openssh-server openssh-client \
        auditd selinux-utils firewalld iptables \
        mariadb-server postgresql redis-server nginx apache2 tomcat9 \
        samba vsftpd nfs-kernel-server; \
    systemctl enable mariadb postgresql redis-server nginx apache2 tomcat9 smbd ssh; \
    apt-get clean; rm -rf /var/lib/apt/lists/*
```

## 文件结构

```
docker/
├── Dockerfile       麒麟靶机镜像
├── systemctl3.py    systemctl 替身脚本（docker-systemctl-replacement）
├── README.md        本说明
└── output/          核查报告输出（docker cp 拷出）
```
