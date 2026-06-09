import paramiko

client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
client.connect('104.207.77.78', port=22022, username='root', password='Ictlbacks77!', timeout=15)

def run(cmd, timeout=30):
    stdin, stdout, stderr = client.exec_command(cmd, timeout=timeout)
    return stdout.read().decode().strip()

print("⛏️ JUNO MINING STATUS")
print("=" * 50)

# Get all accepted/rejected/speed lines
log = run("grep -E 'accepted|speed|height' /opt/junorig/mining.log")

accepted = 0
rejected = 0
last_speed = ""
last_height = ""

for line in log.split('\n'):
    if 'accepted' in line:
        parts = line.split('accepted')[1].strip()
        nums = parts.split(')')[0].replace('(','').split('/')
        accepted = int(nums[0])
        rejected = int(nums[1])
    if 'speed' in line:
        last_speed = line.strip()
    if 'height' in line:
        last_height = line.strip()

print(f"✅ Accepted: {accepted}")
print(f"❌ Rejected: {rejected}")
print(f"📦 {last_height}")
print(f"⚡ {last_speed}")

# Uptime
uptime = run("ps -o etime= -p $(pgrep -f 'junorig -o')")
print(f"⏱️ Uptime: {uptime.strip()}")

# CPU
cpu = run("top -bn1 | grep 'Cpu(s)' | awk '{print $2}'")
print(f"💻 CPU: {cpu}%")

client.close()
