# -*- coding: utf-8 -*-
"""路由層測試：清除資料只能清自己那一份（P0 回歸）。"""

import pytest

import app as A


@pytest.fixture
def client():
    A._DATASETS.clear()
    A._DATASET_ORDER.clear()
    A._TIMES.clear()
    yield A.app.test_client()
    A._DATASETS.clear()
    A._DATASET_ORDER.clear()
    A._TIMES.clear()


def seed(*tokens):
    for t in tokens:
        A._DATASETS[t] = {'fake': t}
        A._DATASET_ORDER.append(t)
        A._TIMES[t] = 0


def test_clear_only_removes_own_token(client):
    seed('tokA', 'tokB')

    assert client.post('/api/clear/tokA').status_code == 200

    assert 'tokA' not in A._DATASETS
    assert 'tokB' in A._DATASETS  # 另一位使用者的資料不受影響
    assert A._DATASET_ORDER == ['tokB']


def test_clear_all_endpoint_is_gone(client):
    """全域清除會誤刪其他人的資料，已移除；不應再對外提供。"""
    seed('tokA', 'tokB')

    assert client.post('/api/clear-all').status_code == 404
    assert set(A._DATASETS) == {'tokA', 'tokB'}


def test_healthz(client):
    r = client.get('/healthz')
    assert r.status_code == 200
    assert r.get_json() == {'ok': True}
