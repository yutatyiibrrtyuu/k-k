"""
换热器设计计算模块
支持两种计算方法：
  1. LMTD 法（对数平均温差法）
  2. NTU-ε 法（传热单元数-效能法）

支持的换热器类型：
  - 顺流（parallel flow）
  - 逆流（counter flow）
  - 交叉流（cross flow，单程，两流体均不混合）
  - 壳管式（shell and tube，1-2 型）
"""

import math
from dataclasses import dataclass
from enum import Enum
from typing import Optional


class HXType(Enum):
    PARALLEL = "parallel"       # 顺流
    COUNTER = "counter"         # 逆流
    CROSS = "cross"             # 交叉流（均不混合）
    SHELL_TUBE = "shell_tube"   # 壳管式 1-2 型


@dataclass
class FluidStream:
    """流体流股"""
    mass_flow: float        # 质量流量 [kg/s]
    cp: float               # 比热容 [J/(kg·K)]
    T_in: float             # 入口温度 [K 或 °C]
    T_out: Optional[float] = None  # 出口温度（已知则填写）

    @property
    def C(self) -> float:
        """热容量流率 [W/K]"""
        return self.mass_flow * self.cp

    @property
    def Q(self) -> Optional[float]:
        """已知进出口温度时的热流量 [W]"""
        if self.T_out is not None:
            return self.C * abs(self.T_out - self.T_in)
        return None


# ---------------------------------------------------------------------------
# LMTD 法
# ---------------------------------------------------------------------------

def lmtd(T_h_in: float, T_h_out: float,
          T_c_in: float, T_c_out: float,
          hx_type: HXType = HXType.COUNTER) -> tuple[float, float]:
    """
    计算对数平均温差及修正系数 F。

    Parameters
    ----------
    T_h_in, T_h_out : 热流体进、出口温度
    T_c_in, T_c_out : 冷流体进、出口温度
    hx_type         : 换热器类型

    Returns
    -------
    (LMTD_corrected, F) : 修正后的对数平均温差 [K]，修正系数 F [-]
    """
    if hx_type == HXType.PARALLEL:
        dT1 = T_h_in - T_c_in
        dT2 = T_h_out - T_c_out
    else:
        # 逆流、交叉流、壳管式均以逆流端差定义
        dT1 = T_h_in - T_c_out
        dT2 = T_h_out - T_c_in

    if dT1 <= 0 or dT2 <= 0:
        raise ValueError(
            f"端差必须为正值，请检查温度设置。dT1={dT1:.2f}, dT2={dT2:.2f}"
        )

    if abs(dT1 - dT2) < 1e-6:
        lmtd_cf = dT1
    else:
        lmtd_cf = (dT1 - dT2) / math.log(dT1 / dT2)

    F = _correction_factor(T_h_in, T_h_out, T_c_in, T_c_out, hx_type)
    return lmtd_cf * F, F


def _correction_factor(T_h_in, T_h_out, T_c_in, T_c_out, hx_type: HXType) -> float:
    """计算 LMTD 修正系数 F（相对逆流对数平均温差）。"""
    if hx_type in (HXType.COUNTER, HXType.PARALLEL):
        return 1.0

    # 无量纲参数
    R = (T_h_in - T_h_out) / (T_c_out - T_c_in) if (T_c_out - T_c_in) != 0 else float("inf")
    P = (T_c_out - T_c_in) / (T_h_in - T_c_in) if (T_h_in - T_c_in) != 0 else 0.0

    if hx_type == HXType.CROSS:
        # Bowman 近似公式（交叉流，两侧均不混合）
        return _F_cross(R, P)
    elif hx_type == HXType.SHELL_TUBE:
        return _F_shell_tube_1_2(R, P)
    return 1.0


def _F_cross(R: float, P: float) -> float:
    """交叉流修正系数（两流体均不混合，Mason 近似）。"""
    # 当 R≈1 时用极限展开
    if abs(R - 1.0) < 1e-4:
        NTU = P / (1 - P) if P < 1 else 50.0
    else:
        arg = (1 - R * P) / (1 - P)
        if arg <= 0:
            return 0.75  # 工程保守估计
        NTU = math.log(arg) / (R - 1)

    if NTU < 1e-6:
        return 1.0

    # 逆流 LMTD（基准）
    dT1_cf = 1 - R * P          # 归一化
    dT2_cf = 1 - P
    if abs(dT1_cf - dT2_cf) < 1e-6:
        lmtd_cf_norm = dT1_cf
    else:
        lmtd_cf_norm = (dT1_cf - dT2_cf) / math.log(dT1_cf / dT2_cf + 1e-12)

    # 交叉流 ε-NTU 关系（两侧均不混合）
    C_ratio = R  # Cmin/Cmax = R（当 R<=1）
    if R <= 1:
        eps = 1 - math.exp(
            (math.exp(-NTU * R * (1 - math.exp(-NTU))) - 1) / R
        ) if NTU > 0 else 0
    else:
        NTU2 = NTU * R
        R2 = 1 / R
        eps_2 = 1 - math.exp(
            (math.exp(-NTU2 * R2 * (1 - math.exp(-NTU2))) - 1) / R2
        ) if NTU2 > 0 else 0
        eps = eps_2 * R  # 转换为热流体侧效能

    Q_norm = eps           # = P (当 R<=1)
    lmtd_cross_norm = Q_norm / NTU if NTU > 0 else lmtd_cf_norm
    F = lmtd_cross_norm / lmtd_cf_norm if lmtd_cf_norm > 0 else 1.0
    return max(0.5, min(1.0, F))


def _F_shell_tube_1_2(R: float, P: float) -> float:
    """壳管式 1-2 型修正系数（标准公式）。"""
    if R == 1.0:
        R += 1e-8
    S = math.sqrt(R**2 + 1)
    denom1 = 2 / P - 1 - R + S
    denom2 = 2 / P - 1 - R - S
    if denom1 <= 0 or denom2 <= 0 or denom1 == denom2:
        return 0.8  # 工程保守估计
    arg = denom1 / denom2
    if arg <= 0:
        return 0.8
    F = (S * math.log((1 - P) / (1 - R * P))) / (
        (R - 1) * math.log(denom1 / denom2)
    ) if (R - 1) != 0 else 1.0
    return max(0.5, min(1.0, F))


def design_lmtd(hot: FluidStream, cold: FluidStream,
                U: float, hx_type: HXType = HXType.COUNTER) -> dict:
    """
    已知换热量，用 LMTD 法求换热面积。

    Parameters
    ----------
    hot  : 热流体（需提供 T_in, T_out）
    cold : 冷流体（需提供 T_in, T_out）
    U    : 总传热系数 [W/(m²·K)]
    hx_type : 换热器类型

    Returns
    -------
    结果字典，含 Q, LMTD, F, A
    """
    Q = hot.C * (hot.T_in - hot.T_out)
    lmtd_val, F = lmtd(hot.T_in, hot.T_out, cold.T_in, cold.T_out, hx_type)
    A = Q / (U * lmtd_val)
    return {
        "Q_W": Q,
        "LMTD_K": lmtd_val,
        "F": F,
        "U_W_m2_K": U,
        "A_m2": A,
        "hx_type": hx_type.value,
    }


# ---------------------------------------------------------------------------
# NTU-ε 法
# ---------------------------------------------------------------------------

def effectiveness(NTU: float, C_ratio: float, hx_type: HXType) -> float:
    """
    由 NTU 和热容比 Cr = Cmin/Cmax 计算效能 ε。

    Parameters
    ----------
    NTU     : 传热单元数
    C_ratio : Cmin / Cmax ∈ [0, 1]
    hx_type : 换热器类型

    Returns
    -------
    ε : 换热效能 ∈ [0, 1]
    """
    Cr = C_ratio
    if Cr < 0 or Cr > 1:
        raise ValueError(f"热容比 C_ratio={Cr:.4f} 超出 [0,1] 范围")

    if hx_type == HXType.PARALLEL:
        return (1 - math.exp(-NTU * (1 + Cr))) / (1 + Cr)

    elif hx_type == HXType.COUNTER:
        if abs(Cr - 1.0) < 1e-6:
            return NTU / (NTU + 1)
        return (1 - math.exp(-NTU * (1 - Cr))) / (1 - Cr * math.exp(-NTU * (1 - Cr)))

    elif hx_type == HXType.CROSS:
        # 两流体均不混合，Kays & London 关联式
        exp1 = math.exp(-Cr * NTU**0.22)
        return 1 - math.exp((exp1 - 1) / (Cr + 1e-12) * NTU**0.78) if Cr > 0 else 1 - math.exp(-NTU)

    elif hx_type == HXType.SHELL_TUBE:
        # 1-2 型壳管式（单壳程双管程）
        if abs(Cr - 1.0) < 1e-6:
            NTU1 = NTU / math.sqrt(2)
            return 2 / (1 + Cr + math.sqrt(2) * (1 + math.exp(-NTU1 * math.sqrt(2))) /
                        (1 - math.exp(-NTU1 * math.sqrt(2))))
        E = math.exp(-NTU * math.sqrt(1 + Cr**2))
        return 2 / (1 + Cr + math.sqrt(1 + Cr**2) * (1 + E) / (1 - E))

    raise ValueError(f"不支持的换热器类型: {hx_type}")


def NTU_from_effectiveness(eps: float, C_ratio: float, hx_type: HXType) -> float:
    """由效能 ε 反推传热单元数 NTU（数值求解）。"""
    if not (0 < eps < 1):
        raise ValueError(f"效能 ε={eps:.4f} 必须在 (0,1) 区间内")

    # 二分法
    lo, hi = 0.0, 500.0
    for _ in range(100):
        mid = (lo + hi) / 2
        if effectiveness(mid, C_ratio, hx_type) < eps:
            lo = mid
        else:
            hi = mid
        if hi - lo < 1e-9:
            break
    return (lo + hi) / 2


def design_ntu(hot: FluidStream, cold: FluidStream,
               A: float, U: float, hx_type: HXType = HXType.COUNTER) -> dict:
    """
    已知换热面积，用 NTU-ε 法求出口温度和换热量。

    Parameters
    ----------
    hot, cold : 流体（仅需提供 T_in，T_out 可为 None）
    A         : 换热面积 [m²]
    U         : 总传热系数 [W/(m²·K)]
    hx_type   : 换热器类型

    Returns
    -------
    结果字典，含 ε, NTU, Q, 以及两股流体出口温度
    """
    C_hot = hot.C
    C_cold = cold.C
    C_min = min(C_hot, C_cold)
    C_max = max(C_hot, C_cold)
    C_ratio = C_min / C_max

    NTU = U * A / C_min
    eps = effectiveness(NTU, C_ratio, hx_type)

    Q = eps * C_min * (hot.T_in - cold.T_in)
    T_hot_out = hot.T_in - Q / C_hot
    T_cold_out = cold.T_in + Q / C_cold

    return {
        "NTU": NTU,
        "C_ratio": C_ratio,
        "epsilon": eps,
        "Q_W": Q,
        "T_hot_out_K": T_hot_out,
        "T_cold_out_K": T_cold_out,
        "hx_type": hx_type.value,
    }


def rate_ntu(hot: FluidStream, cold: FluidStream,
             U: float, hx_type: HXType = HXType.COUNTER) -> dict:
    """
    已知进出口温度，用 NTU-ε 法求所需换热面积。

    Parameters
    ----------
    hot, cold : 流体（T_in 和 T_out 均需提供）
    U         : 总传热系数 [W/(m²·K)]

    Returns
    -------
    结果字典
    """
    if hot.T_out is None or cold.T_out is None:
        raise ValueError("rate_ntu 需要提供进出口温度")

    C_hot = hot.C
    C_cold = cold.C
    C_min = min(C_hot, C_cold)
    C_max = max(C_hot, C_cold)
    C_ratio = C_min / C_max

    Q = C_hot * (hot.T_in - hot.T_out)
    Q_max = C_min * (hot.T_in - cold.T_in)
    eps = Q / Q_max

    NTU = NTU_from_effectiveness(eps, C_ratio, hx_type)
    A = NTU * C_min / U

    return {
        "NTU": NTU,
        "C_ratio": C_ratio,
        "epsilon": eps,
        "Q_W": Q,
        "U_W_m2_K": U,
        "A_m2": A,
        "hx_type": hx_type.value,
    }


# ---------------------------------------------------------------------------
# 污垢热阻 & 总传热系数
# ---------------------------------------------------------------------------

def overall_U(h_hot: float, h_cold: float,
              R_f_hot: float = 0.0, R_f_cold: float = 0.0,
              t_wall: float = 0.002, k_wall: float = 50.0) -> float:
    """
    计算总传热系数 U（平壁近似）。

    Parameters
    ----------
    h_hot, h_cold : 热侧/冷侧对流换热系数 [W/(m²·K)]
    R_f_hot, R_f_cold : 热侧/冷侧污垢热阻 [(m²·K)/W]
    t_wall        : 壁厚 [m]
    k_wall        : 壁面导热系数 [W/(m·K)]

    Returns
    -------
    U [W/(m²·K)]
    """
    R_total = (1 / h_hot) + R_f_hot + (t_wall / k_wall) + R_f_cold + (1 / h_cold)
    return 1 / R_total


# ---------------------------------------------------------------------------
# 结果打印
# ---------------------------------------------------------------------------

def print_result(result: dict, title: str = "换热器计算结果") -> None:
    """格式化打印计算结果。"""
    labels = {
        "Q_W": ("换热量 Q", "W", 2),
        "LMTD_K": ("对数平均温差 LMTD", "K", 4),
        "F": ("LMTD 修正系数 F", "-", 4),
        "U_W_m2_K": ("总传热系数 U", "W/(m²·K)", 2),
        "A_m2": ("换热面积 A", "m²", 4),
        "NTU": ("传热单元数 NTU", "-", 4),
        "C_ratio": ("热容比 Cr", "-", 4),
        "epsilon": ("效能 ε", "-", 4),
        "T_hot_out_K": ("热流体出口温度", "K", 2),
        "T_cold_out_K": ("冷流体出口温度", "K", 2),
        "hx_type": ("换热器类型", "", 0),
    }
    print(f"\n{'='*50}")
    print(f"  {title}")
    print(f"{'='*50}")
    for key, val in result.items():
        if key in labels:
            name, unit, prec = labels[key]
            if isinstance(val, float):
                print(f"  {name:<28} {val:.{prec}f} {unit}")
            else:
                print(f"  {name:<28} {val} {unit}")
    print(f"{'='*50}\n")


# ---------------------------------------------------------------------------
# 示例
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("  换热器计算示例")
    print("=" * 60)

    # ------------------------------------------------------------------
    # 示例 1：LMTD 法设计计算（逆流）
    # 热水从 80°C 冷却到 50°C，冷水从 20°C 加热到 45°C
    # ------------------------------------------------------------------
    print("\n【示例 1】LMTD 法 — 逆流换热器设计")
    hot1 = FluidStream(mass_flow=2.0, cp=4182.0, T_in=80.0, T_out=50.0)
    cold1 = FluidStream(mass_flow=3.0, cp=4182.0, T_in=20.0, T_out=45.0)
    U1 = overall_U(h_hot=5000, h_cold=3000, R_f_hot=1e-4, R_f_cold=2e-4)
    res1 = design_lmtd(hot1, cold1, U=U1, hx_type=HXType.COUNTER)
    print_result(res1, "逆流换热器 LMTD 法")

    # ------------------------------------------------------------------
    # 示例 2：NTU-ε 法校核计算
    # 已知面积 A=4 m²，求出口温度
    # ------------------------------------------------------------------
    print("【示例 2】NTU-ε 法 — 校核计算（逆流）")
    hot2 = FluidStream(mass_flow=2.0, cp=4182.0, T_in=353.15)  # 80°C
    cold2 = FluidStream(mass_flow=3.0, cp=4182.0, T_in=293.15)  # 20°C
    res2 = design_ntu(hot2, cold2, A=4.0, U=U1, hx_type=HXType.COUNTER)
    print_result(res2, "逆流换热器 NTU-ε 法（校核）")

    # ------------------------------------------------------------------
    # 示例 3：壳管式换热器（1-2型）
    # ------------------------------------------------------------------
    print("【示例 3】壳管式换热器（1-2 型）NTU-ε 设计")
    hot3 = FluidStream(mass_flow=1.5, cp=2100.0, T_in=150.0, T_out=90.0)
    cold3 = FluidStream(mass_flow=2.0, cp=4182.0, T_in=25.0, T_out=60.0)
    U3 = overall_U(h_hot=800, h_cold=3000)
    res3 = rate_ntu(hot3, cold3, U=U3, hx_type=HXType.SHELL_TUBE)
    print_result(res3, "壳管式换热器 NTU-ε 法（设计）")

    # ------------------------------------------------------------------
    # 示例 4：交叉流换热器
    # ------------------------------------------------------------------
    print("【示例 4】交叉流换热器 NTU-ε 设计")
    hot4 = FluidStream(mass_flow=0.5, cp=1005.0, T_in=200.0, T_out=80.0)   # 空气
    cold4 = FluidStream(mass_flow=0.3, cp=4182.0, T_in=15.0, T_out=70.0)   # 水
    U4 = overall_U(h_hot=150, h_cold=3000)
    res4 = rate_ntu(hot4, cold4, U=U4, hx_type=HXType.CROSS)
    print_result(res4, "交叉流换热器 NTU-ε 法（设计）")
