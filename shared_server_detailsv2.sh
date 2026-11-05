#!/usr/bin/env bash
# ACRA – Shared server details for the weekly / monthly monitoring report.
#
# Run on the Linux server (cPanel / CloudLinux or any systemd Linux) as root:
#
#     bash shared_server_details.sh
#
# It collects the two report sections, Server Overview and Service Health, into one Excel
# workbook laid out like the report, puts it in the web server's document root
# (/var/www/html) under a long random name and prints the download link. Upload the file
# on the report page in ACRA (Shared reports: "Load server details from Excel").
#
# READ-ONLY. The script only reads: /proc, /etc/os-release, uname, df, ip, systemctl and
# the rpm / dpkg package databases. It installs nothing, starts / stops / restarts
# nothing and changes no setting. Its only network connection is the kernel update check:
# dnf / yum / apt download the repository lists (as "dnf check-update" or "apt update" would)
# into the script's private temporary folder, not into the system's package cache (see
# "kernel update check" below; --offline skips it). It writes only the Excel file (built in
# the private temporary folder that is removed at the end) and deletes the published file
# again after --keep-minutes.
#
# Security of the download:
#   * The file name is 32 random hex characters (not guessable) and an existing file is
#     never overwritten.
#   * It is deleted automatically after 30 minutes (--keep-minutes, 0 = keep it).
#   * The file shows the hostname, IPs, kernel and running services: treat it as
#     confidential. If the document root has no index page, directory listing may show
#     the file to anyone – the script warns you. The link is plain http.
#   * Most private: --no-web keeps the file in the current folder (mode 600) instead;
#     copy it with SFTP / WinSCP.
#
# Options:
#   --docroot DIR        where to put the file (default /var/www/html)
#   --keep-minutes N     delete the file after N minutes (default 30, 0 = keep, max 1440)
#   --url-host HOST      host name / IP for the printed link (default: the server's public IP)
#   --no-web             don't publish; write ./<hostname>_server_details.xlsx (mode 600)
#   --offline            don't download the repository lists for the kernel check; use only
#                        the lists already on the server (they may be out of date)
#
# Needs bash and one of: zip, python3 or python (only to pack the .xlsx file).
set -u
export LC_ALL=C LANG=C SYSTEMD_PAGER= PAGER=cat
umask 077

MARKER="ACRA Shared server details"
VERSION=2
DOCROOT=/var/www/html
KEEP=30
URL_HOST=
NO_WEB=0
OFFLINE=0

die() { echo "$*" >&2; exit 1; }
have() { command -v "$1" >/dev/null 2>&1; }

while [ $# -gt 0 ]; do
  case "$1" in
    --docroot) DOCROOT="${2:-}"; shift 2 ;;
    --keep-minutes) KEEP="${2:-}"; shift 2 ;;
    --url-host) URL_HOST="${2:-}"; shift 2 ;;
    --no-web) NO_WEB=1; shift ;;
    --offline) OFFLINE=1; shift ;;
    --check-repos) shift ;;   # v1 option; the repositories are now always checked
    -h|--help) sed -n '2,40p' "$0"; exit 0 ;;
    *) die "Unknown option: $1 (see --help)" ;;
  esac
done
[[ "$KEEP" =~ ^[0-9]+$ ]] && [ "$KEEP" -le 1440 ] || die "--keep-minutes must be 0 to 1440"
[ -z "$URL_HOST" ] || [[ "$URL_HOST" =~ ^[A-Za-z0-9.:\[\]-]{1,253}$ ]] || die "--url-host: give a host name or IP only"

PACKER=
if have zip; then PACKER=zip
elif have python3; then PACKER=python3
elif have python; then PACKER=python
else die "Needs zip, python3 or python to make the .xlsx file (none found)."; fi

TMP=$(mktemp -d) || die "Couldn't make a temporary folder"
trap 'rm -rf "$TMP"' EXIT

echo "Collecting... (read-only)"

# ---------------------------------------------------------------- helpers
human() {   # bytes -> 7.61 GiB
  awk -v n="$1" 'BEGIN{split("B KiB MiB GiB TiB",u," ");i=1;while(n>=1024&&i<5){n/=1024;i++}
    if(i==1)printf "%d %s",n,u[i]; else printf "%.2f %s",n,u[i]}'
}
xml() {     # text -> safe XML cell text (no control characters, never a formula)
  local s
  s=$(printf '%s' "$1" | tr -d '\000-\010\013\014\016-\037' | cut -c1-500)
  case "$s" in
    [=+@]*) s="'$s" ;;
    -*) [[ "$s" =~ ^-[0-9] ]] || s="'$s" ;;
  esac
  printf '%s' "$s" | sed -e 's/&/\&amp;/g' -e 's/</\&lt;/g' -e 's/>/\&gt;/g' -e 's/"/\&quot;/g'
}
col() { local n=$1 s=""; while [ "$n" -gt 0 ]; do n=$((n-1)); s=$(printf "\\$(printf '%03o' $((65 + n % 26)))")$s; n=$((n / 26)); done; printf '%s' "$s"; }

# row FILE ROWNO "style|text" ...   (style: 0 plain, 1 title, 2 header, 3 label, 4 note)
row() {
  local f=$1 r=$2 c=0 cell st tx; shift 2
  printf '<row r="%d">' "$r" >> "$f"
  for cell in "$@"; do
    c=$((c+1)); st=${cell%%|*}; tx=${cell#*|}
    [ -z "$tx" ] && [ "$st" = 0 ] && continue
    printf '<c r="%s%d" t="inlineStr" s="%s"><is><t xml:space="preserve">%s</t></is></c>' \
      "$(col $c)" "$r" "$st" "$(xml "$tx")" >> "$f"
  done
  printf '</row>' >> "$f"
}
sheet_start() {   # FILE width...
  local f=$1 i=0 w; shift
  printf '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' > "$f"
  for w in "$@"; do i=$((i+1)); printf '<col min="%d" max="%d" width="%s" customWidth="1"/>' $i $i "$w" >> "$f"; done
  printf '</cols><sheetData>' >> "$f"
}
sheet_end() { printf '</sheetData></worksheet>' >> "$1"; }

# ---------------------------------------------------------------- Server Overview
HOST=$(hostname -f 2>/dev/null || hostname 2>/dev/null || uname -n)
OS=""
if [ -r /etc/os-release ]; then OS=$(. /etc/os-release 2>/dev/null; echo "${PRETTY_NAME:-}"); fi
[ -n "$OS" ] || OS=$(cat /etc/redhat-release 2>/dev/null || true)
CL=No
{ [ -e /etc/cloudlinux-release ] || grep -qi cloudlinux /etc/os-release 2>/dev/null; } && CL=Yes
LOAD=$(awk '{printf "%s, %s, %s", $1, $2, $3}' /proc/loadavg 2>/dev/null)
KERNEL=$(uname -r)

# global addresses of the real interfaces (not Docker / veth / lo); public IPv4 first
IPS=""
if have ip; then
  IPS=$(ip -o addr show scope global 2>/dev/null | awk '$2 !~ /^(docker|br-|veth|virbr|cni|flannel|cali|tun|tap|lo)/ {split($4,a,"/"); print a[1]}' |
    awk '!seen[$0]++' | awk '{
      p = ($0 ~ /:/) ? 2 : ($0 ~ /^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.|100\.(6[4-9]|[7-9][0-9]|1[01][0-9]|12[0-7])\.)/ ? 1 : 0);
      print p "\t" $0}' | sort -s -k1,1n | cut -f2 | paste -sd, - | sed 's/,/, /g')
fi

# ---------------------------------------------------------------- kernel update check
# Asks the package repositories whether a newer kernel exists, the same answer as a manual
# "dnf check-update kernel" / "apt update && apt list --upgradable", but read-only:
#   * the repository lists are downloaded into this script's private temporary folder
#     (removed at the end), never into /var/cache/dnf, /var/cache/yum or /var/lib/apt;
#   * dnf's log, history and "countme" folders point at the temporary folder too;
#   * apt runs with its update hooks switched off (they would write notifier / stamp files);
#   * nothing is installed, upgraded or removed, and no package database is changed.
# yum (CentOS / CloudLinux 7) keeps its usual short-lived lock file (/var/run/yum.pid) while it
# runs. --offline skips the download and uses only the lists already on the server.
newest_of() { grep -v '^$' | sort -V | tail -1; }
newer_than_running() {   # versions on stdin -> the newest one if it is newer than the running kernel
  local newest
  newest=$( (cat; echo "$KERNEL") | newest_of)
  [ "$newest" != "$KERNEL" ] && echo "$newest"
}
KU=""; KNOTE=""; KNEW=""; KLIVE=""
if have rpm; then
  ARCH=$(uname -m)
  INST=$(rpm -q --qf '%{VERSION}-%{RELEASE}.%{ARCH}\n' kernel kernel-core 2>/dev/null | grep -v 'not installed' | newer_than_running || true)
  REPO=""; SRC=""
  PM=$(command -v dnf || command -v yum || true)
  if [ -n "$PM" ]; then
    PMN=$(basename "$PM")
    mkdir -p "$TMP/pm/cache" "$TMP/pm/log" "$TMP/pm/persist"
    OPTS=(-q -y --disableexcludes=all --setopt=persistdir="$TMP/pm/persist")
    [ "$PMN" = dnf ] && OPTS+=(--setopt=logdir="$TMP/pm/log")
    LIST=(list --showduplicates available kernel kernel-core)
    [ "$PMN" = dnf ] && LIST=(list --showduplicates --available kernel kernel-core)
    ok=1
    if [ "$OFFLINE" = 0 ]; then
      echo "  Checking the package repositories for a newer kernel (downloads the lists to a temporary folder)..."
      MC=(makecache); [ "$PMN" = yum ] && MC=(makecache fast)
      if timeout 900 "$PM" "${OPTS[@]}" --setopt=cachedir="$TMP/pm/cache" "${MC[@]}" >/dev/null 2>"$TMP/pm/err"; then
        OUTL=$(timeout 300 "$PM" "${OPTS[@]}" --setopt=cachedir="$TMP/pm/cache" -C "${LIST[@]}" 2>/dev/null); ok=0
        SRC="the package repositories, checked $(date -u '+%Y-%m-%d %H:%M UTC')"
      else
        echo "  Couldn't reach the repositories; using the server's own $PMN cache instead."
      fi
    fi
    if [ "$ok" != 0 ]; then
      OUTL=$(timeout 300 "$PM" "${OPTS[@]}" -C "${LIST[@]}" 2>/dev/null); rc=$?
      # rc 1 with no output = nothing newer listed OR no usable cache: only trust a cache that exists
      if [ $rc = 0 ] || ls /var/cache/"$PMN"/ 2>/dev/null | grep -q .; then ok=0; SRC="the server's own $PMN cache, which may be out of date"; fi
    fi
    if [ "$ok" = 0 ]; then
      REPO=$(printf '%s\n' "$OUTL" | awk -v a="$ARCH" '$1 ~ /^kernel(-core)?\./ && NF >= 2 {
          split($1, n, "."); v = $2; sub(/^[0-9]+:/, "", v); print v "." n[length(n)] }' |
        grep -E "\.($ARCH|noarch)$" | newer_than_running || true)
    fi
  fi
  if [ -n "$REPO" ] && [ "$REPO" != "$INST" ] && [ "$(printf '%s\n%s\n' "$REPO" "$INST" | newest_of)" = "$REPO" ]; then
    KU=Yes; KNEW=$REPO; KNOTE="Kernel $REPO is available in the repositories and not installed yet"
    [ -n "$INST" ] && KNOTE="$KNOTE; kernel $INST is already installed and waits for a reboot"
    KNOTE="$KNOTE (from $SRC)"
  elif [ -n "$INST" ]; then
    KU=Yes; KNEW=$INST; KNOTE="Newer kernel $INST is installed; a reboot is needed to use it"
  elif [ -n "$SRC" ]; then
    KU=No; KNEW="$KERNEL (running)"; KNOTE="Running the newest kernel (from $SRC)"
  else
    KNOTE="Couldn't check the repositories (no dnf / yum, or no network and no local cache); choose Yes / No in ACRA"
  fi
elif have dpkg-query && have apt-get; then
  # The server's apt settings are copied into the temporary folder with the update hooks
  # cleared (#clear); the lists go to the temporary folder; no package cache file is written.
  A="$TMP/apt"; mkdir -p "$A/parts" "$A/lists/partial" "$A/cache/archives/partial"
  cp -p /etc/apt/apt.conf.d/* "$A/parts/" 2>/dev/null || true
  { cat /etc/apt/apt.conf 2>/dev/null; printf '\n'; } > "$A/main.conf"
  CLR='#clear APT::Update::Pre-Invoke;\n#clear APT::Update::Post-Invoke;\n#clear APT::Update::Post-Invoke-Success;\n'
  printf "$CLR" > "$A/parts/zzzz-acra-read-only"; printf "$CLR" >> "$A/main.conf"
  printf 'Dir::Etc::parts "%s/parts";\nDir::Etc::main "%s/main.conf";\n' "$A" "$A" > "$A/apt.conf"
  AO=(-o Dir::Cache="$A/cache" -o Dir::Cache::pkgcache= -o Dir::Cache::srcpkgcache= -o APT::Sandbox::User=root)
  LISTS=/var/lib/apt/lists; SRC=""
  if [ "$OFFLINE" = 0 ]; then
    echo "  Checking the package repositories for a newer kernel (downloads the lists to a temporary folder)..."
    if APT_CONFIG="$A/apt.conf" timeout 900 apt-get -qq "${AO[@]}" -o Dir::State::Lists="$A/lists" update >/dev/null 2>"$A/err" &&
       ! grep -q '^[WE]: Failed to fetch' "$A/err"; then
      LISTS="$A/lists"; SRC="the package repositories, checked $(date -u '+%Y-%m-%d %H:%M UTC')"
    else
      echo "  Couldn't reach all repositories; using the server's own apt lists instead."
    fi
  fi
  [ -n "$SRC" ] || SRC="the server's own apt lists, which may be out of date"
  U=$(APT_CONFIG="$A/apt.conf" apt list --upgradable "${AO[@]}" -o Dir::State::Lists="$LISTS" 2>/dev/null |
      grep '^linux-image-' | awk '{split($1, n, "/"); print n[1] " " $2}' | head -1)
  N=$(ls /boot 2>/dev/null | sed -n 's/^vmlinuz-//p' | newer_than_running || true)
  if [ -n "$U" ]; then
    KU=Yes; KNEW=$U; KNOTE="Kernel update available: $U (from $SRC)"
    [ -n "$N" ] && KNOTE="$KNOTE; kernel $N is already installed and waits for a reboot"
  elif [ -n "$N" ]; then KU=Yes; KNEW=$N; KNOTE="Newer kernel $N is installed; a reboot is needed to use it"
  else KU=No; KNEW="$KERNEL (running)"; KNOTE="Running the newest kernel (from $SRC)"; fi
else
  KNOTE="Couldn't check (no rpm or dpkg)"
fi
# live patching (KernelCare / TuxCare on CloudLinux): the patched kernel level, read locally
if have kcarectl; then
  KLIVE=$(timeout 30 kcarectl --uname 2>/dev/null | head -1)
  [ -n "$KLIVE" ] && KLIVE="KernelCare: running kernel is patched to the level of $KLIVE"
fi

# file systems (no pseudo / CageFS / container mounts)
df -P -T -k 2>/dev/null | awk 'NR>1 && NF>=7 {
    m=$7; for(i=8;i<=NF;i++) m=m" "$i
    if ($2 ~ /^(tmpfs|devtmpfs|overlay|squashfs|proc|sysfs|cgroup2?|nsfs|autofs|debugfs|tracefs|securityfs|mqueue|hugetlbfs|pstore|bpf|configfs|fusectl|binfmt_misc|efivarfs|ramfs|iso9660|devpts|rpc_pipefs|fuse\..*)$/) next
    if (m ~ /^\/(proc|sys|dev|run|snap)(\/|$)|cagefs|\/var\/lib\/docker\/|\/var\/lib\/containers\//) next
    if (seen[$1]++ || $3 <= 0) next
    print m "\t" $4*1024 "\t" $5*1024 "\t" $3*1024 "\t" (($4+$5)>0 ? sprintf("%.0f%%", $4/($4+$5)*100) : "0%")
  }' > "$TMP/fs.tsv"

MEM=$(awk '/^(MemTotal|MemAvailable|MemFree|Buffers|Cached|SwapTotal|SwapFree):/ {v[$1]=$2*1024}
  END { if (!("MemAvailable:" in v)) v["MemAvailable:"]=v["MemFree:"]+v["Buffers:"]+v["Cached:"]
        printf "%s %s %s %s", v["MemTotal:"]+0, v["MemAvailable:"]+0, v["SwapTotal:"]+0, v["SwapFree:"]+0 }' /proc/meminfo)
read -r MT MA ST SF <<< "$MEM"

# ---------------------------------------------------------------- Service Health
friendly() {   # unit name -> "order|readable name" (known services are listed first)
  case "$1" in
    httpd) echo "01|Apache Web Server (httpd)" ;; apache2) echo "01|Apache Web Server" ;;
    lshttpd|lsws|litespeed) echo "02|LiteSpeed Web Server" ;; nginx) echo "03|Nginx Web Server" ;;
    ea-nginx) echo "03|Nginx (cPanel)" ;;
    mysqld|mysql) echo "10|MySQL Database" ;; mariadb) echo "10|MariaDB Database" ;;
    postgresql*) echo "11|PostgreSQL Database" ;; mongod) echo "12|MongoDB Database" ;;
    cpanel) echo "20|cPanel Service" ;; cpsrvd) echo "20|cPanel Service (cpsrvd)" ;;
    cpanel_php_fpm) echo "21|cPanel PHP-FPM" ;; tailwatchd) echo "22|cPanel TailWatch" ;;
    exim) echo "30|Exim Mail Server" ;; postfix) echo "30|Postfix Mail Server" ;;
    dovecot) echo "31|Dovecot (IMAP / POP3)" ;; spamd|spamassassin) echo "32|SpamAssassin" ;;
    named) echo "40|DNS Server (BIND)" ;; pdns) echo "40|DNS Server (PowerDNS)" ;;
    pure-ftpd) echo "50|FTP Server (Pure-FTPd)" ;; proftpd) echo "50|FTP Server (ProFTPD)" ;;
    vsftpd) echo "50|FTP Server (vsftpd)" ;;
    sshd|ssh) echo "60|SSH Service" ;; crond|cron) echo "61|Cron Service" ;;
    lfd) echo "70|CSF Login Failure Daemon (lfd)" ;; csf) echo "70|ConfigServer Firewall (csf)" ;;
    firewalld) echo "71|Firewall (firewalld)" ;; fail2ban) echo "72|Fail2ban" ;;
    imunify360*|imunify-antivirus|aibolit-resident) echo "73|Imunify360" ;;
    clamd*) echo "74|ClamAV Antivirus" ;;
    ea-php*-php-fpm|php-fpm|php*-fpm) echo "80|PHP-FPM" ;;
    redis*) echo "81|Redis" ;; memcached) echo "82|Memcached" ;; varnish) echo "83|Varnish Cache" ;;
    haproxy) echo "84|HAProxy" ;; tomcat*) echo "85|Apache Tomcat" ;; docker) echo "86|Docker" ;;
    elasticsearch) echo "87|Elasticsearch" ;; zabbix-agent|zabbix-agent2) echo "88|Zabbix Agent" ;;
    kcare|kernelcare) echo "89|KernelCare" ;; cagefs) echo "90|CageFS" ;; lvestats) echo "91|LVE Stats" ;;
    db_governor) echo "92|MySQL Governor" ;;
    *) echo "" ;;
  esac
}
# system plumbing that says nothing about the server's health in a client report
SKIP='^(systemd-.*|dbus.*|getty@.*|serial-getty@.*|user@.*|user-runtime-dir@.*|polkit|rsyslog|syslog|auditd|chronyd|chrony|ntpd|irqbalance|tuned|NetworkManager.*|network.*|wpa_supplicant|udisks2|accounts-daemon|rtkit-daemon|ModemManager|multipathd|lvm2-.*|dm-event|qemu-guest-agent|hv-kvp-daemon|hv_.*|waagent|walinuxagent|google-.*|amazon-ssm-agent|snap.*|packagekit|unattended-upgrades|atd|kdump|smartd|mdmonitor|gssproxy|rpcbind|nfs-.*|sssd.*|oddjobd|abrt.*|cloud-.*|emergency|rescue|plymouth.*|kmod-.*|console-.*|blk-availability|import-state|loadmodules|ldconfig|selinux-.*|rngd|mcelog|microcode|iscsi.*|lm_sensors|cups.*|avahi.*|bluetooth|alsa-.*|switcheroo.*|fwupd|thermald|upower|colord|gdm|lightdm)$'

: > "$TMP/svc.tsv"
if have systemctl; then
  systemctl list-units --type=service --all --no-legend --no-pager --plain 2>/dev/null |
  while read -r unit load active sub desc; do
    case "$unit" in *.service) ;; *) continue ;; esac
    [ "$load" = loaded ] || continue
    name=${unit%.service}
    fr=$(friendly "$name")
    if [ "$active" = failed ]; then :                                 # always shown
    elif [ -n "$fr" ]; then
      if [ "$active" = inactive ] && [ "$sub" = dead ]; then
        # a known service that is installed but stopped: shown only if it is enabled
        [ "$(systemctl is-enabled "$unit" 2>/dev/null)" = enabled ] || continue
      elif [ "$active" != active ]; then continue; fi
    else
      [ "$active" = active ] && [ "$sub" = running ] || continue
      printf '%s' "$name" | grep -Eq "$SKIP" && continue
    fi
    case "$sub" in running) st=Running ;; exited) st="Active (exited)" ;; dead) st=Stopped ;; *) st=${sub^} ;; esac
    [ "$active" = failed ] && st=Failed
    printf '%s\t%s\t%s\t%s\t%s\t%s\n' "${fr%%|*}" "$( [ -n "$fr" ] && echo "${fr#*|}" || echo "$name")" "$st" "$unit" "$name" "$desc"
  done > "$TMP/picked.tsv"
  # Active Since for every picked unit (one systemctl call per 50 units)
  cut -f4 "$TMP/picked.tsv" | xargs -r -n 50 systemctl show --no-pager -p Id -p ActiveEnterTimestamp 2>/dev/null |
    awk -F= '/^Id=/{id=$2} /^ActiveEnterTimestamp=/{ if (match($2, /[0-9]{4}-[0-9]{2}-[0-9]{2}/)) print id "\t" substr($2, RSTART, 10); else print id "\t" }' > "$TMP/since.tsv"
  awk -F'\t' 'NR==FNR {s[$1]=$2; next} { o = ($1 == "" ? "999" : $1); d = ($3 == "Running" || $3 == "Active (exited)") ? s[$4] : "";
      print o "\t" $2 "\t" $3 "\t" d "\t" $5 "\t" $6 }' "$TMP/since.tsv" "$TMP/picked.tsv" |
    sort -t$'\t' -k1,1n -k2,2f | cut -f2- | tr '\t' '\037' > "$TMP/svc.tsv"   # \037: empty fields survive read
fi

running_for() {
  [ -n "$1" ] || return 0
  local d; d=$(( ( $(date -u +%s) - $(date -u -d "$1" +%s 2>/dev/null || date -u +%s) ) / 86400 ))
  [ "$d" -ge 0 ] && { [ "$d" = 1 ] && echo "1 day" || echo "$d days"; }
}

# ---------------------------------------------------------------- the Excel file
COLLECTED=$(date -u '+%Y-%m-%d %H:%M UTC')
X="$TMP/x"; mkdir -p "$X/_rels" "$X/xl/_rels" "$X/xl/worksheets"

S1="$X/xl/worksheets/sheet1.xml"; sheet_start "$S1" 28 34 60 14 10; r=0
row "$S1" $((r+=1)) "1|Server Overview"
row "$S1" $((r+=1)) "4|$MARKER · collected $COLLECTED"
r=$((r+1))
row "$S1" $((r+=1)) "2|Category" "2|Details" "2|"
row "$S1" $((r+=1)) "3|Server Hostname" "0|$HOST"
row "$S1" $((r+=1)) "3|Server IP Address" "0|$IPS"
row "$S1" $((r+=1)) "3|Server OS Version" "0|$OS"
row "$S1" $((r+=1)) "3|CloudLinux Server" "0|$CL"
row "$S1" $((r+=1)) "3|Server Load Average" "0|$LOAD"
row "$S1" $((r+=1)) "3|Current Kernel Version" "0|$KERNEL"
row "$S1" $((r+=1)) "3|Kernel Update Available" "0|$KU" "4|$KNOTE"
row "$S1" $((r+=1)) "3|Newest Kernel Available" "0|$KNEW"
[ -n "$KLIVE" ] && row "$S1" $((r+=1)) "3|Kernel Live Patching" "0|$KLIVE"
r=$((r+1))
row "$S1" $((r+=1)) "1|File System, Memory and Swap"
row "$S1" $((r+=1)) "2|Partition" "2|Used" "2|Free" "2|Total" "2|Used %"
NFS=0
while IFS=$'\t' read -r m u f t p; do
  row "$S1" $((r+=1)) "3|$m" "0|$(human "$u")" "0|$(human "$f")" "0|$(human "$t")" "0|$p"; NFS=$((NFS+1))
done < "$TMP/fs.tsv"
[ "$NFS" -gt 0 ] || row "$S1" $((r+=1)) "0|(no file systems found)"
r=$((r+1))
row "$S1" $((r+=1)) "2|Memory" "2|Total" "2|Used" "2|Free" "2|Used %"
pct() { awk -v u="$1" -v t="$2" 'BEGIN{printf "%.0f%%", (t>0 ? u/t*100 : 0)}'; }
[ "${MT:-0}" -gt 0 ] && row "$S1" $((r+=1)) "3|Memory (RAM)" "0|$(human "$MT")" "0|$(human $((MT-MA)))" "0|$(human "$MA")" "0|$(pct $((MT-MA)) "$MT")"
row "$S1" $((r+=1)) "3|Swap" "0|$(human "${ST:-0}")" "0|$(human $((ST-SF)))" "0|$(human "${SF:-0}")" "0|$(pct $((ST-SF)) "$ST")"
sheet_end "$S1"

S2="$X/xl/worksheets/sheet2.xml"; sheet_start "$S2" 34 16 14 13 28 50; r=0
row "$S2" $((r+=1)) "1|Service Health"
row "$S2" $((r+=1)) "4|$MARKER · collected $COLLECTED"
r=$((r+1))
row "$S2" $((r+=1)) "2|Service" "2|Status" "2|Active Since" "2|Running For" "2|Unit" "2|Description"
NSVC=0
while IFS=$'\037' read -r svc st since unit desc; do
  row "$S2" $((r+=1)) "3|$svc" "0|$st" "0|$since" "0|$(running_for "$since")" "0|$unit" "0|$desc"; NSVC=$((NSVC+1))
done < "$TMP/svc.tsv"
sheet_end "$S2"

S3="$X/xl/worksheets/sheet3.xml"; sheet_start "$S3" 18 60; r=0
row "$S3" $((r+=1)) "1|About this file"; r=$((r+1))
row "$S3" $((r+=1)) "3|Generator" "0|$MARKER"
row "$S3" $((r+=1)) "3|Version" "0|$VERSION"
row "$S3" $((r+=1)) "3|Hostname" "0|$HOST"
row "$S3" $((r+=1)) "3|Collected (UTC)" "0|$COLLECTED"; r=$((r+1))
row "$S3" $((r+=1)) "4|Upload this file on the Shared report page in ACRA. Remove the services you don't want in the report there – nothing is changed on the server."
sheet_end "$S3"

cat > "$X/[Content_Types].xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
EOF
cat > "$X/_rels/.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
EOF
cat > "$X/xl/workbook.xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Server Overview" sheetId="1" r:id="rId1"/><sheet name="Service Health" sheetId="2" r:id="rId2"/><sheet name="About" sheetId="3" r:id="rId3"/></sheets></workbook>
EOF
cat > "$X/xl/_rels/workbook.xml.rels" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
EOF
cat > "$X/xl/styles.xml" <<'EOF'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FF2E75B6"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><i/><sz val="9"/><color rgb="FF666666"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDEEAF6"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs>
<cellXfs count="5"><xf borderId="1" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>
<xf fontId="1" applyFont="1"/>
<xf fontId="2" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf fontId="0" fillId="3" borderId="1" applyFill="1" applyBorder="1"><alignment vertical="top"/></xf>
<xf fontId="3" applyFont="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
EOF

OUT="$TMP/out.xlsx"
case "$PACKER" in
  zip) (cd "$X" && zip -q -X -r "$OUT" '[Content_Types].xml' _rels xl) ;;
  *)   (cd "$X" && "$PACKER" -c 'import os, sys, zipfile
z = zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED)
for top in ("[Content_Types].xml", "_rels", "xl"):
    for root, dirs, files in os.walk(top) if os.path.isdir(top) else [(".", [], [top])]:
        for f in files:
            p = os.path.normpath(os.path.join(root, f))
            z.write(p, p.replace(os.sep, "/"))
z.close()' "$OUT") ;;
esac
[ -s "$OUT" ] || die "Couldn't make the Excel file."
echo "  $NFS file systems, $NSVC services"

# ---------------------------------------------------------------- publishing
if [ "$NO_WEB" = 1 ]; then
  short=$(printf '%s' "${HOST%%.*}" | tr -c 'A-Za-z0-9._-' '_')
  DEST="$PWD/${short:-server}_server_details.xlsx"
  [ -e "$DEST" ] || [ -L "$DEST" ] && die "$DEST already exists – move or delete it first."
  ( set -C; cat "$OUT" > "$DEST" ) || die "Couldn't write $DEST"
  chmod 600 "$DEST"
  echo; echo "Saved (only you can read it): $DEST"; echo "Copy it with SFTP / WinSCP and upload it in ACRA."
  exit 0
fi

[ -d "$DOCROOT" ] || die "$DOCROOT doesn't exist. Use --docroot DIR, or --no-web to keep the file here."
DOCROOT=$(cd "$DOCROOT" && pwd -P)
NAME="$(od -An -N16 -tx1 /dev/urandom | tr -d ' \n').xlsx"
DEST="$DOCROOT/$NAME"
( umask 022; set -C; cat "$OUT" > "$DEST" ) 2>/dev/null || die "Can't write to $DOCROOT – run as root (sudo), or use --no-web."
chmod 644 "$DEST"        # the web server must be able to read it

HOSTLINK=$URL_HOST
if [ -z "$HOSTLINK" ]; then
  HOSTLINK=$(printf '%s' "$IPS" | tr ',' '\n' | tr -d ' ' | grep -v ':' |
    grep -Ev '^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.|100\.(6[4-9]|[7-9][0-9]|1[01][0-9]|12[0-7])\.)' | head -1)
  [ -n "$HOSTLINK" ] || HOSTLINK=$(printf '%s' "$IPS" | cut -d, -f1 | tr -d ' ')
  [ -n "$HOSTLINK" ] || HOSTLINK=$HOST
fi
case "$HOSTLINK" in *:*) [[ "$HOSTLINK" == \[* ]] || HOSTLINK="[$HOSTLINK]" ;; esac

echo; echo "Download link:"; echo "  http://$HOSTLINK/$NAME"; echo
if ! ls "$DOCROOT" 2>/dev/null | grep -Eqi '^index\.(html?|php|shtml)$'; then
  echo "WARNING: $DOCROOT has no index page. If directory listing is on, anyone opening"
  echo "http://$HOSTLINK/ can see this file. Download it now; it is removed automatically."
fi
if [ "$KEEP" -gt 0 ]; then
  # a small detached process removes the file later (survives logout)
  if have setsid; then RUNNER="setsid"; else RUNNER=""; fi
  $RUNNER nohup sh -c 'sleep "$1"; [ -f "$2" ] && [ ! -L "$2" ] && rm -f -- "$2"' _ $((KEEP * 60)) "$DEST" >/dev/null 2>&1 &
  [ "$KEEP" = 1 ] && echo "The file is deleted automatically in 1 minute." || echo "The file is deleted automatically in $KEEP minutes."
fi
echo "Remove it now (after downloading):  rm -f $DEST"
echo "The file holds the hostname, IPs, kernel and services – keep it confidential."
