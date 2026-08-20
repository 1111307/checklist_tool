# =====================================================================
#  配置核查 · 麒麟靶机镜像
#  用途：本地用 Docker 起一台「银河麒麟 V10」靶机，用于自测
#        kylin/check_kylin.sh 的核查效果（无需真实麒麟服务器）。
#
#  基础镜像：macrosan/kylin（银河麒麟高级服务器 V10，rpm/yum 系）
#  组件：systemd + MariaDB + PostgreSQL + Redis + Nginx + Apache + Tomcat
#        + Samba + vsftpd + NFS + auditd + firewalld + SSH
#  说明：达梦(DM8)、人大金仓(KingbaseES) 为商业数据库，需授权文件，
#        本镜像不内置，请用现成镜像单独起容器（见 README）。
#
#  构建（在项目根目录执行，构建上下文 = 整个项目）：
#    docker build -f docker/Dockerfile -t kylin-target:latest .
#
#  运行：
#    docker run -d --privileged --name kylin-target \
#      -v /sys/fs/cgroup:/sys/fs/cgroup:ro \
#      kylin-target:latest
#
#  进入并跑核查脚本：
#    docker exec -it kylin-target bash -c "bash /opt/check_kylin.sh"
# =====================================================================

FROM macrosan/kylin:v10-sp3-2403

LABEL maintainer="配置核查项目"
LABEL description="配置核查靶机：银河麒麟V10 + 数据库 + 中间件 + 文件共享 + systemd"

ENV TZ=Asia/Shanghai \
    LANG=zh_CN.UTF-8 \
    container=docker

# systemctl 替身作为 PID 1 时，用 SIGTERM 优雅停机
STOPSIGNAL SIGTERM

RUN set -eux; \
    # ---------- 基础工具 + systemd + 安全/审计组件 ----------
    yum install -y --setopt=tsflags=nodocs --setopt=install_weak_deps=False \
        systemd systemd-libs dbus polkit sudo passwd iproute net-tools hostname \
        which procps-ng cronie rsyslog logrotate util-linux util-linux-user \
        openssh-server openssh-clients openssl \
        audit audit-libs policycoreutils selinux-policy-targeted \
        firewalld iptables-services ebtables python3 \
    ; \
    # ---------- 数据库 ----------
    yum install -y --setopt=tsflags=nodocs --setopt=install_weak_deps=False \
        mariadb-server mariadb \
        postgresql-server postgresql \
    ; \
    # ---------- 中间件 ----------
    yum install -y --setopt=tsflags=nodocs --setopt=install_weak_deps=False \
        redis nginx httpd tomcat \
    ; \
    # ---------- 文件共享 ----------
    yum install -y --setopt=tsflags=nodocs --setopt=install_weak_deps=False \
        samba samba-client vsftpd nfs-utils \
    ; \
    # ---------- 初始化数据库（失败不阻断，systemd 首次启动也会自动初始化） ----------
    ( mysql_install_db --user=mysql --datadir=/var/lib/mysql >/dev/null 2>&1 || \
      mariadb-install-db --user=mysql --datadir=/var/lib/mysql >/dev/null 2>&1 ) || true; \
    ( postgresql-setup --initdb >/dev/null 2>&1 || \
      initdb -D /var/lib/pgsql/data >/dev/null 2>&1 ) || true; \
    # ---------- 精简 systemd 单元以适配容器 ----------
    ( cd /lib/systemd/system/sysinit.target.wants/; \
      for i in *; do [ "$i" = systemd-tmpfiles-setup.service ] || rm -f "$i"; done ); \
    rm -f /lib/systemd/system/multi-user.target.wants/*; \
    rm -f /etc/systemd/system/*.wants/*; \
    rm -f /lib/systemd/system/local-fs.target.wants/*; \
    rm -f /lib/systemd/system/sockets.target.wants/*udev*; \
    rm -f /lib/systemd/system/sockets.target.wants/*initctl*; \
    rm -f /lib/systemd/system/basic.target.wants/*; \
    rm -f /lib/systemd/system/anaconda.target.wants/*; \
    # ---------- 启用服务（容器以 systemd 启动时自动拉起） ----------
    systemctl enable mariadb postgresql redis nginx httpd tomcat smb vsftpd sshd \
        >/dev/null 2>&1 || true; \
    # ---------- 清理缓存减小体积 ----------
    yum clean all; \
    rm -rf /var/cache/yum/* /var/log/yum.log

# 拷入核查脚本
COPY kylin/check_kylin.sh /opt/check_kylin.sh
RUN chmod +x /opt/check_kylin.sh

# =====================================================================
#  systemctl 替身（docker-systemctl-replacement）
#  用 Python 脚本替换 systemctl，在没有 systemd 的环境下也能执行
#  systemctl is-active / is-enabled / start / stop 等命令。
#  适用于 Docker Desktop / WSL2（cgroup v2 与 Kylin systemd 243
#  的 cgroup v1 不兼容，无法真跑 systemd）。
#  容器以该脚本作为 PID 1 时，自动进入 init 模式并拉起 enabled 服务。
# =====================================================================
COPY docker/systemctl3.py /usr/local/bin/systemctl3.py
RUN set -eux; \
    chmod +x /usr/local/bin/systemctl3.py; \
    mv /usr/bin/systemctl /usr/bin/systemctl.real 2>/dev/null || true; \
    ln -sf /usr/local/bin/systemctl3.py /usr/bin/systemctl; \
    ln -sf /usr/local/bin/systemctl3.py /bin/systemctl 2>/dev/null || true; \
    /usr/bin/systemctl --version

VOLUME [ "/sys/fs/cgroup" ]

# 以 systemctl 替身作为 PID 1（init 模式，自动拉起 enabled 服务）
CMD [ "/usr/bin/systemctl" ]
