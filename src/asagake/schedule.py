from typing import List

def plan_batches(symbols: List[str], batch_size: int = 50) -> List[List[str]]:
    """kabu API の登録上限 50 に合わせたバッチ分割"""
    return [symbols[i:i+batch_size] for i in range(0, len(symbols), batch_size)]
