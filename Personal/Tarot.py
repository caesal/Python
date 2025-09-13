import random
from typing import List, Dict, Tuple

# ===== 牌面定义 =====
MAJOR_ARCANA: List[str] = [
    "The Fool", "The Magician", "The High Priestess", "The Empress",
    "The Emperor", "The Hierophant", "The Lovers", "The Chariot",
    "Strength", "The Hermit", "Wheel of Fortune", "Justice",
    "The Hanged Man", "Death", "Temperance", "The Devil",
    "The Tower", "The Star", "The Moon", "The Sun",
    "Judgement", "The World",
]

SUITS: List[str] = ["Wands", "Cups", "Swords", "Pentacles"]
RANKS: List[str] = [str(n) for n in range(1, 11)] + ["J", "Q", "K"]
MINOR_ARCANA: List[str] = [f"{rank} of {suit}" for suit in SUITS for rank in RANKS]
ALL_CARDS: List[str] = MAJOR_ARCANA + MINOR_ARCANA

# ===== 英文牌名到中文牌名映射 =====
MAJOR_CN: Dict[str, str] = {
    "The Fool": "愚者", "The Magician": "魔术师", "The High Priestess": "女祭司",
    "The Empress": "皇后", "The Emperor": "皇帝", "The Hierophant": "教皇",
    "The Lovers": "恋人", "The Chariot": "战车", "Strength": "力量",
    "The Hermit": "隐者", "Wheel of Fortune": "命运之轮", "Justice": "正义",
    "The Hanged Man": "倒吊人", "Death": "死神", "Temperance": "节制",
    "The Devil": "恶魔", "The Tower": "高塔", "The Star": "星星",
    "The Moon": "月亮", "The Sun": "太阳", "Judgement": "审判", "The World": "世界",
}
SUIT_CN = {"Wands": "权杖", "Cups": "圣杯", "Swords": "宝剑", "Pentacles": "星币"}
NUM_CN = {
    "1": "一", "2": "二", "3": "三", "4": "四", "5": "五",
    "6": "六", "7": "七", "8": "八", "9": "九", "10": "十",
    "J": "侍者", "Q": "王后", "K": "国王",
}

def cn_card_name(card: str) -> str:
    if card in MAJOR_CN:
        return MAJOR_CN[card]
    rank, _, suit = card.partition(" of ")
    suit_cn = SUIT_CN.get(suit, suit)
    rank_cn = NUM_CN.get(rank, rank)
    return f"{suit_cn}{rank_cn}"

def display_name(card: str) -> str:
    return f"{card}（{cn_card_name(card)}）"

# ===== 中文含义字典 正位用原键 逆位在键名后加 _rev =====
MEANINGS: Dict[str, str] = {
    # 大阿尔卡纳
    "The Fool": "正位 启程 自由 纯真 相信直觉 勇敢迈步",
    "The Fool_rev": "逆位 冲动 犯险 困惑 不切实际 方向不明",
    "The Magician": "正位 意志 专注 资源整合 行动力 创造力",
    "The Magician_rev": "逆位 技能错用 操控 欺瞒 意志涣散 目标混乱",
    "The High Priestess": "正位 直觉 潜意识 静观 保密 内在智慧",
    "The High Priestess_rev": "逆位 直觉受阻 隐情外泄 疑心 逃避内在声音",
    "The Empress": "正位 丰饶 照护 美感 关系滋养 项目开花",
    "The Empress_rev": "逆位 过度依赖 停滞 过度纵容 创造力受阻",
    "The Emperor": "正位 结构 规则 责任 领导 稳固基础",
    "The Emperor_rev": "逆位 刚愎 独断 失控 缺乏边界 责任缺失",
    "The Hierophant": "正位 传统 学习 导师 制度 正规路径",
    "The Hierophant_rev": "逆位 墨守成规 盲从 形式空洞 非主流尝试",
    "The Lovers": "正位 选择 共鸣 亲密 价值一致 心之契合",
    "The Lovers_rev": "逆位 分歧 诱惑 摇摆 价值冲突 关系考验",
    "The Chariot": "正位 决心 驾驭 前进 胜利 意志统一",
    "The Chariot_rev": "逆位 失衡 偏执 漫无目标 控制失效",
    "Strength": "正位 温柔的力量 勇气 自律 包容 坚韧",
    "Strength_rev": "逆位 自我怀疑 冲动 压抑 怯懦 力量内耗",
    "The Hermit": "正位 独处 省察 智者之光 寻求真义",
    "The Hermit_rev": "逆位 封闭 逃避 孤立 无法落实洞见",
    "Wheel of Fortune": "正位 转机 循环 机遇 因果 顺势而为",
    "Wheel of Fortune_rev": "逆位 阻滞 倒霉 错失时机 抗拒变化",
    "Justice": "正位 公平 真相 平衡 因果 负责选择",
    "Justice_rev": "逆位 失衡 偏颇 不公 逃避承担 自欺",
    "The Hanged Man": "正位 视角转换 暂停 牺牲 以退为进",
    "The Hanged Man_rev": "逆位 固执 拖延 无谓牺牲 停滞不前",
    "Death": "正位 结束与重生 断舍离 更新 转化",
    "Death_rev": "逆位 抗拒改变 旧习难改 拖延 无法放下",
    "Temperance": "正位 节制 调和 渐进 平衡 顺其自然",
    "Temperance_rev": "逆位 失衡 过度 极端 焦躁 难以协调",
    "The Devil": "正位 诱惑 束缚 执念 欲望 合同代价",
    "The Devil_rev": "逆位 松绑 觉醒 戒除 解脱 面对阴影",
    "The Tower": "正位 崩解 真相闪电 旧结构瓦解 突破",
    "The Tower_rev": "逆位 迟迟不崩 更大代价 侥幸 暗涌",
    "The Star": "正位 希望 灵感 复原 信任 宇宙回应",
    "The Star_rev": "逆位 迷茫 幻灭 自我否定 光芒受遮",
    "The Moon": "正位 潜意识 梦境 直觉不安 朦胧 需要等待",
    "The Moon_rev": "逆位 烦扰散去 真相渐明 情绪稳定",
    "The Sun": "正位 喜悦 清晰 成功 童心 活力盛放",
    "The Sun_rev": "逆位 延迟的小确幸 自信不足 乐观被遮",
    "Judgement": "正位 召唤 觉醒 听从心声 重启与清算",
    "Judgement_rev": "逆位 自我怀疑 拖延 错过召唤 审视不足",
    "The World": "正位 完成 圆满 整合 里程碑 新阶段开启",
    "The World_rev": "逆位 未竟之事 循环未闭 缺一角 需要完善",

    # 权杖 Wands
    "1 of Wands": "正位 灵感点火 新计划 行动力 开始执行",
    "1 of Wands_rev": "逆位 灵感受阻 犹豫 不敢启动",
    "2 of Wands": "正位 规划 远景 选择版图 勇于拓展",
    "2 of Wands_rev": "逆位 眼界受限 迟疑 错过窗口",
    "3 of Wands": "正位 发展在路上 合作 出海 期待回响",
    "3 of Wands_rev": "逆位 配合不良 进展拖延 资源不顺",
    "4 of Wands": "正位 阶段性庆祝 稳定 基础落地 家与圈层",
    "4 of Wands_rev": "逆位 仪式走样 稳定感动摇",
    "5 of Wands": "正位 竞争 磨合 碰撞中学习 试炼",
    "5 of Wands_rev": "逆位 内耗 纷争升级 无意义争执",
    "6 of Wands": "正位 胜利 认可 领先 光环时刻",
    "6 of Wands_rev": "逆位 期待落空 自负 招致非议",
    "7 of Wands": "正位 立场坚定 防守反击 捍卫成果",
    "7 of Wands_rev": "逆位 防线松动 妥协 放弃据点",
    "8 of Wands": "正位 加速 突破 高效沟通 快速推进",
    "8 of Wands_rev": "逆位 讯息延迟 阻力 方向混乱",
    "9 of Wands": "正位 坚守 临门一脚 戒备但不放弃",
    "9 of Wands_rev": "逆位 透支 疲惫 戒心过重 想放弃",
    "10 of Wands": "正位 负担 责任压顶 但可以扛过",
    "10 of Wands_rev": "逆位 放下不必要负担 委派 精简",
    "J of Wands": "正位 信使 新火花 学习与尝试 勇敢表达",
    "J of Wands_rev": "逆位 三分钟热度 浮躁 信息不实",
    "Q of Wands": "正位 自信 领导魅力 热情点燃他人",
    "Q of Wands_rev": "逆位 焦躁 嫉妒 控制欲强",
    "K of Wands": "正位 远见 统筹 以身作则 定方向",
    "K of Wands_rev": "逆位 专断 冒进 承诺过度",

    # 圣杯 Cups
    "1 of Cups": "正位 情感涌现 关怀 灵感温柔 新关系契机",
    "1 of Cups_rev": "逆位 情绪淤积 表达受阻 空杯需要自我滋养",
    "2 of Cups": "正位 互相吸引 互信 合作共鸣",
    "2 of Cups_rev": "逆位 失衡 付出不对等 分歧",
    "3 of Cups": "正位 友谊 庆祝 社群支持 共享喜悦",
    "3 of Cups_rev": "逆位 过度社交 八卦 疏离",
    "4 of Cups": "正位 倦怠 冷感 内观 重新评估",
    "4 of Cups_rev": "逆位 苏醒 接受机会 情绪回流",
    "5 of Cups": "正位 失落 悲伤 聚焦已失 忽略仍在",
    "5 of Cups_rev": "逆位 走出哀伤 修复 看见剩余",
    "6 of Cups": "正位 旧识 怀旧 纯真 善意回访",
    "6 of Cups_rev": "逆位 沉溺过去 依赖 迟滞成长",
    "7 of Cups": "正位 选择繁多 幻象与愿景 需要筛选",
    "7 of Cups_rev": "逆位 看清现实 定下抉择 聚焦",
    "8 of Cups": "正位 转身离开 寻找更高价值",
    "8 of Cups_rev": "逆位 难舍难分 犹豫 不愿放手",
    "9 of Cups": "正位 心愿达成 满足 感恩 自在",
    "9 of Cups_rev": "逆位 满足成空 浮夸 过度享乐",
    "10 of Cups": "正位 圆满和谐 家与心的联结",
    "10 of Cups_rev": "逆位 表面和谐 内在裂痕 期望落差",
    "J of Cups": "正位 灵感使者 温柔讯息 艺术心",
    "J of Cups_rev": "逆位 过敏 多愁 情绪摇摆",
    "Q of Cups": "正位 共情 直觉深刻 治愈能量",
    "Q of Cups_rev": "逆位 情绪被动 过度牺牲 自我忽略",
    "K of Cups": "正位 情绪成熟 稳定包容 智慧表达",
    "K of Cups_rev": "逆位 情感操控 压抑或冷漠 不坦诚",

    # 宝剑 Swords
    "1 of Swords": "正位 灵光清明 真相 理性开端",
    "1 of Swords_rev": "逆位 迷雾 误解 迟迟不决",
    "2 of Swords": "正位 两难 权衡 需要摘下眼罩",
    "2 of Swords_rev": "逆位 决策偏差 拖延 自欺",
    "3 of Swords": "正位 心伤 分离 锋利的真相",
    "3 of Swords_rev": "逆位 疗愈 缝合 释怀",
    "4 of Swords": "正位 休整 静思 暂停以养锋",
    "4 of Swords_rev": "逆位 过度停滞 焦虑 难以休息",
    "5 of Swords": "正位 争执 赢得表面 输掉关系",
    "5 of Swords_rev": "逆位 和解 反思 知所进退",
    "6 of Swords": "正位 渡河 过渡 远离纷扰",
    "6 of Swords_rev": "逆位 迁移受阻 情绪拉扯",
    "7 of Swords": "正位 策略 隐秘 运筹 取巧",
    "7 of Swords_rev": "逆位 真相暴露 归还 面对错误",
    "8 of Swords": "正位 自限 困缚 恐惧使然",
    "8 of Swords_rev": "逆位 解放 走出困局 自我解绑",
    "9 of Swords": "正位 焦虑 失眠 过度担忧",
    "9 of Swords_rev": "逆位 担忧缓解 光线透入",
    "10 of Swords": "正位 结束 最糟已过 放下重启",
    "10 of Swords_rev": "逆位 迟来的结束 拔刀见阳",
    "J of Swords": "正位 信息侦察 学习敏锐 直言不讳",
    "J of Swords_rev": "逆位 冲动言语 浅尝辄止 好辩",
    "Q of Swords": "正位 清醒独立 辩证公允 设立边界",
    "Q of Swords_rev": "逆位 刻薄 偏见 冷硬",
    "K of Swords": "正位 逻辑 权威 决断 公正执行",
    "K of Swords_rev": "逆位 刚愎 冷酷 滥用权力",

    # 星币 Pentacles
    "1 of Pentacles": "正位 实际机会 种子 资源到位",
    "1 of Pentacles_rev": "逆位 机会流失 预算受限 土壤未备",
    "2 of Pentacles": "正位 平衡多线 弹性 调度有术",
    "2 of Pentacles_rev": "逆位 失衡 忙乱 超载",
    "3 of Pentacles": "正位 团队协作 工艺 专业认可",
    "3 of Pentacles_rev": "逆位 配合不畅 标准不清 成果打折",
    "4 of Pentacles": "正位 持有 稳健 设界 降低风险",
    "4 of Pentacles_rev": "逆位 过度保守 执着所有 缺乏流动",
    "5 of Pentacles": "正位 困难期 资源匮乏 但仍可求援",
    "5 of Pentacles_rev": "逆位 走出低谷 支援到来",
    "6 of Pentacles": "正位 给予与接受 慷慨 资源流动",
    "6 of Pentacles_rev": "逆位 失衡交易 附带条件 施与受不均",
    "7 of Pentacles": "正位 耕耘 评估 等待收成",
    "7 of Pentacles_rev": "逆位 心急 收获寡 方向需要调整",
    "8 of Pentacles": "正位 打磨技艺 专注 迭代精进",
    "8 of Pentacles_rev": "逆位 机械重复 烦倦 品质下降",
    "9 of Pentacles": "正位 自主 自给自足 享受成果",
    "9 of Pentacles_rev": "逆位 过度依赖 虚荣 外强中干",
    "10 of Pentacles": "正位 家业 财富传承 集体稳定",
    "10 of Pentacles_rev": "逆位 家族纷争 结构性风险",
    "J of Pentacles": "正位 学徒心 新任务 脚踏实地",
    "J of Pentacles_rev": "逆位 拖延 三心二意 缺乏耐心",
    "Q of Pentacles": "正位 照护 实用 慷慨而稳",
    "Q of Pentacles_rev": "逆位 过度操心 控物忽人 溺爱或苛责",
    "K of Pentacles": "正位 成就 稳固 经营有方",
    "K of Pentacles_rev": "逆位 守旧 固化 贪婪或安全感缺失",
}

def meaning_for(card: str, orientation: str) -> str:
    key = card if orientation == "Upright" else f"{card}_rev"
    return MEANINGS.get(key, "暂未录入该牌的中文释义")

class TarotDeck:
    def __init__(self) -> None:
        self.reset()

    def reset(self) -> None:
        self.cards = ALL_CARDS.copy()
        random.shuffle(self.cards)

    def draw(self, n: int = 1) -> List[Tuple[str, str]]:
        if n < 1:
            raise ValueError("至少抽取一张")
        if n > len(self.cards):
            raise ValueError("牌库不足")
        drawn: List[Tuple[str, str]] = []
        for _ in range(n):
            card = self.cards.pop()
            orientation = random.choice(["Upright", "Reversed"])
            drawn.append((card, orientation))
        return drawn

def print_reading(drawn: List[Tuple[str, str]], exp = True) -> None:
    # 第一部分 只列牌
    print("本次抽牌")
    for idx, (card, orientation) in enumerate(drawn, start=1):
        ori_cn = "正位" if orientation == "Upright" else "逆位"
        print(f"{idx}. {display_name(card)} {ori_cn}")
    print()

    if exp:
        # 第二部分 只列释义
        print("释义汇总")
        for idx, (card, orientation) in enumerate(drawn, start=1):
            ori_cn = "正位" if orientation == "Upright" else "逆位"
            print(f"{idx}. {display_name(card)} {ori_cn}")
            print("   " + meaning_for(card, orientation))
            print()

if __name__ == "__main__":
    deck = TarotDeck()
    picks = deck.draw(3)
    print_reading(picks,0) #exp = 1 to bring the brief explanation back
