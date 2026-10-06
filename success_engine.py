# -*- coding: utf-8 -*-
"""
减肥成功率引擎（精简版）
========================
模型：第 i 天净消耗 X_i ~ N(mu_i, sigma_i^2)，各天独立。
累计净消耗 X = Σ X_i ~ N(mu_sum, sigma_sum^2)，
其中 mu_sum = Σ mu_i, sigma_sum^2 = Σ sigma_i^2。
由中心极限定理，天数足够时可用正态近似。

统一用「标准化 + 标准正态 CDF Φ」：
    P(X >= T) = 1 - Φ((T - mu_sum) / sigma_sum)
"""
import math

try:
    from scipy.stats import norm
    def _phi(x):
        return norm.cdf(x)              # 标准正态 CDF
except ImportError:
    def _phi(x):
        return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def success_rate_engine(daily, target_kg, last_month=None):
    """
    参数
    ----
    daily : list[tuple]
        每日 (predicted_net_calories_i, error_std_i)，单位千卡
    target_kg : float
        目标减重公斤数（1kg ≈ 7700 千卡）
    last_month : tuple[float, float] | None
        上月累计的 (mu_last, sigma_last)（累计净消耗均值、误差标准差）

    返回
    ----
    dict：final_prob / half_prob / improvement_prob / conf_interval
    """
    mu_sum = sum(mu for mu, _ in daily)
    sigma_sum = math.sqrt(sum(s * s for _, s in daily))

    def clip(x):
        return round(min(1.0, max(0.0, x)), 3)

    def prob_ge(T):
        """P(X >= T) = 1 - Φ((T - mu_sum)/sigma_sum)"""
        if sigma_sum == 0:
            return 1.0 if mu_sum >= T else 0.0
        return 1.0 - _phi((T - mu_sum) / sigma_sum)

    result = {}

    # 1) 达标成功率：T = 7700 * target_kg
    result['final_prob'] = clip(prob_ge(7700 * target_kg))

    # 2) 减半目标成功率：T = 7700 * target_kg / 2
    result['half_prob'] = clip(prob_ge(7700 * target_kg / 2))

    # 3) 比上月进步率：D = X_本月 - X_上月
    #    mu_D = mu_sum - mu_last, sigma_D^2 = sigma_sum^2 + sigma_last^2
    #    P(D >= 0) = Φ(mu_D / sigma_D)
    if last_month is not None:
        mu_last, sigma_last = last_month
        mu_D = mu_sum - mu_last
        sigma_D = math.sqrt(sigma_sum ** 2 + sigma_last ** 2)
        result['improvement_prob'] = (
            clip(1.0 if mu_D >= 0 else 0.0) if sigma_D == 0
            else clip(_phi(mu_D / sigma_D))
        )
    else:
        result['improvement_prob'] = None

    # 4) 置信区间：mu_sum ± z*sigma_sum, z=1.645/1.96/2.576
    zs = {'90%': 1.645, '95%': 1.96, '99%': 2.576}
    result['conf_interval'] = {
        k: (round(mu_sum - z * sigma_sum, 3), round(mu_sum + z * sigma_sum, 3))
        for k, z in zs.items()
    }

    return result


if __name__ == '__main__':
    # 30 天，每天 μ=300、σ=100，目标减 2kg；上月累计 (7500, 600)
    daily = [(300, 100)] * 30
    r = success_rate_engine(daily, target_kg=2, last_month=(7500, 600))
    print('输入：30天，每天 μ=300, σ=100；目标 2kg；上月累计(7500, 600)')
    for k, v in r.items():
        print(f'{k}: {v}')
