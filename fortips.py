# -*- coding: utf-8 -*-
# 解析fortigate 6.X的conf,將IP,group,policy匯出成excel 方便審查
import sys
import re
import ipaddress
from openpyxl import Workbook

def parse_address(lines):
    addr_list = []
    current_name = None

    for line in lines:
        line = line.strip()

        match_edit = re.match(r'^edit\s+"([^"]+)"$', line)
        if match_edit:
            current_name = match_edit.group(1)
            continue

        match_subnet = re.match(r'^set subnet ([\d\.]+) ([\d\.]+)', line)
        if match_subnet and current_name:
            ip = match_subnet.group(1)
            mask = match_subnet.group(2)
            try:
                net = ipaddress.IPv4Network(f"{ip}/{mask}", strict=False)
                ip_only = str(net.network_address)
                addr_list.append((current_name, ip_only))
            except ValueError:
                addr_list.append((current_name, ip))
            continue

        if line == 'next':
            current_name = None

    return addr_list

def parse_addrgrp(lines, addr_dict):
    groups = []
    current_group = None

    for line in lines:
        line = line.strip()

        match_edit = re.match(r'^edit\s+"([^"]+)"$', line)
        if match_edit:
            current_group = match_edit.group(1)
            continue

        match_member = re.match(r'^set member (.+)$', line)
        if match_member and current_group:
            members = re.findall(r'"([^"]+)"', match_member.group(1))
            for m in members:
                ip = addr_dict.get(m, 'N/A')
                groups.append((current_group, m, ip))
            continue

        if line == 'next':
            current_group = None

    return groups

def parse_policy(lines):
    policies = []
    current_policy = {}

    for line in lines:
        line = line.strip()

        if line.startswith('edit '):
            if current_policy:
                policies.append(current_policy)
            current_policy = {'id': line.split(' ')[1]}
            continue

        if line.startswith('set name '):
            current_policy['name'] = re.findall(r'"(.*?)"', line)[0]

        if line.startswith('set srcintf '):
            current_policy['srcintf'] = re.findall(r'"(.*?)"', line)

        if line.startswith('set dstintf '):
            current_policy['dstintf'] = re.findall(r'"(.*?)"', line)

        if line.startswith('set srcaddr '):
            current_policy['srcaddr'] = re.findall(r'"(.*?)"', line)

        if line.startswith('set dstaddr '):
            current_policy['dstaddr'] = re.findall(r'"(.*?)"', line)

        if line.startswith('set service '):
            current_policy['service'] = re.findall(r'"(.*?)"', line)

        if line.startswith('set action '):
            current_policy['action'] = line.split(' ')[2]

        if line.startswith('set status '):
            current_policy['status'] = line.split(' ')[2]

        if line == 'next':
            continue

    if current_policy:
        policies.append(current_policy)

    return policies

def main(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    addr_lines = []
    addrgrp_lines = []
    policy_lines = []
    in_addr = False
    in_addrgrp = False
    in_policy = False

    for line in lines:
        if line.strip() == 'config firewall address':
            in_addr = True
            in_addrgrp = False
            in_policy = False
            continue
        elif line.strip() == 'config firewall addrgrp':
            in_addr = False
            in_addrgrp = True
            in_policy = False
            continue
        elif line.strip() == 'config firewall policy':
            in_addr = False
            in_addrgrp = False
            in_policy = True
            continue
        elif line.strip() == 'end':
            in_addr = False
            in_addrgrp = False
            in_policy = False
            continue

        if in_addr:
            addr_lines.append(line)
        elif in_addrgrp:
            addrgrp_lines.append(line)
        elif in_policy:
            policy_lines.append(line)

    addr_list = parse_address(addr_lines)
    addr_dict = {name: ip for name, ip in addr_list}
    group_list = parse_addrgrp(addrgrp_lines, addr_dict)
    policy_list = parse_policy(policy_lines)

    wb = Workbook()
    wb.remove(wb.active)

    # 活頁 address
    ws_addr = wb.create_sheet(title='address')
    ws_addr.append(['名稱', 'IP'])
    for name, ip in addr_list:
        ws_addr.append([name, ip])

    # 活頁 group
    ws_group = wb.create_sheet(title='group')
    ws_group.append(['群組名稱', '成員', 'IP'])
    for group, member, ip in group_list:
        ws_group.append([group, member, ip])

    # 活頁 policy
    ws_policy = wb.create_sheet(title='policy')
    ws_policy.append(['ID', '名稱', '來源介面', '目的介面', '來源地址', '目的地址', '服務', '動作', '狀態'])
    for p in policy_list:
        ws_policy.append([
            p.get('id', ''),
            p.get('name', ''),
            ", ".join(p.get('srcintf', [])),
            ", ".join(p.get('dstintf', [])),
            ", ".join(p.get('srcaddr', [])),
            ", ".join(p.get('dstaddr', [])),
            ", ".join(p.get('service', [])),
            p.get('action', ''),
            p.get('status', '')
        ])

    wb.save(output_file)
    print(f"[✓] 完成產生: {output_file}")

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("用法: python fortips.py 整包設定檔.conf 輸出.xlsx")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
