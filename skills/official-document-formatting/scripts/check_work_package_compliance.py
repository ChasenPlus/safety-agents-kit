import os

base = os.path.expandvars(r'$SAK_KB_ROOT\03-输出层\工作包\金隅兴发科技园三期分布式光伏')

# D组检查
d_checks = {
    'S-02_高处作业专项施工方案.md': ['1.2m', '180mm', '安全带', '5级风', '6级风', '高挂低用', '工具袋', '抛掷', '1.5m', '1/20'],
    'S-03_起重吊装专项施工方案.md': ['十不吊', '200~300mm', '45°', '80%', '5级风', '6级风', '吊物下', '100kN', '10t'],
    'S-10_临时用电专项施工方案.md': ['30mA', '0.1s', '4Ω', '2.5m', '4m', '6m', '一机一闸', 'TN-S', '三级配电'],
    'S-17_动火作业专项施工方案.md': ['10m', '5m', '企安安', '5级风', '监护人', '30分钟', '动火票'],
    'S-13_光伏并网调试专项施工方案.md': ['NB/T 32042', '停电验电', '防拉弧', '72小时', '直流保险', '触电'],
}
for fname, checks in d_checks.items():
    fpath = os.path.join(base, 'D组专项方案', fname)
    with open(fpath, 'r', encoding='utf-8') as f:
        text = f.read()
    print('\n===== ' + fname + ' =====')
    for chk in checks:
        ok = chk in text
        status = 'OK' if ok else 'MISS'
        print('  [' + status + '] ' + chk)

# 岗位工作包结构检查
print('\n===== 岗位工作包完整性抽查 =====')
posts = ['施工项目经理', '专职安全员', '焊工', '电工', '屋面安装工']
for p in posts:
    fpath = os.path.join(base, '岗位工作包', '金隅兴发屋顶光伏板_' + p + '_安全交底.md')
    with open(fpath, 'r', encoding='utf-8') as f:
        text = f.read()
    has_risk = '危险源' in text
    has_measure = '安全技术措施' in text
    has_conf = '交底确认' in text
    print(p + ': 危险源=' + str(has_risk) + ' 措施=' + str(has_measure) + ' 确认=' + str(has_conf) + ' 长度=' + str(len(text)))

# 检查各岗位交底内容长度是否合理（应有一定内容，非空壳）
print('\n===== 岗位工作包内容深度检查 =====')
for p in ['施工项目经理', '项目技术负责人', '专职安全员', '质检员', '资料员', '汽车吊司机', '起重司索信号工', '电工', '焊工', '屋面安装班组长', '屋面安装工']:
    lens = {}
    for doc_type in ['安全交底', '技术交底', '应急处置卡', '考核标准', '制度规范卡', '施工方案']:
        fpath = os.path.join(base, '岗位工作包', '金隅兴发屋顶光伏板_' + p + '_' + doc_type + '.md')
        if os.path.exists(fpath):
            with open(fpath, 'r', encoding='utf-8') as f:
                lens[doc_type] = len(f.read())
        else:
            lens[doc_type] = -1
    print(p + ': ' + str(lens))
