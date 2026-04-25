"""
API 测试脚本
"""
import requests
import json
from pathlib import Path

BASE_URL = "http://localhost:8000"


def test_health_check():
    """测试健康检查"""
    print("测试健康检查...")
    response = requests.get(f"{BASE_URL}/health")
    print(f"状态码: {response.status_code}")
    print(f"响应: {response.json()}")
    print()


def test_upload_files():
    """测试文件上传"""
    print("测试文件上传...")

    # 准备测试文件（需要实际的文件）
    files = [
        ('files', ('test1.xlsx', open('test1.xlsx', 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')),
    ]

    response = requests.post(f"{BASE_URL}/api/v1/shipments/upload", files=files)
    print(f"状态码: {response.status_code}")
    print(f"响应: {json.dumps(response.json(), indent=2, ensure_ascii=False)}")

    return response.json()


def test_process_shipment(shipment_id):
    """测试开始处理"""
    print(f"测试开始处理票据 {shipment_id}...")

    data = {
        "exclude_anti_dumping": False,
        "min_similarity": 0.6
    }

    response = requests.post(
        f"{BASE_URL}/api/v1/shipments/{shipment_id}/process",
        json=data
    )
    print(f"状态码: {response.status_code}")
    print(f"响应: {json.dumps(response.json(), indent=2, ensure_ascii=False)}")

    return response.json()


def test_task_status(task_id):
    """测试查询任务状态"""
    print(f"测试查询任务状态 {task_id}...")

    response = requests.get(f"{BASE_URL}/api/v1/tasks/{task_id}/status")
    print(f"状态码: {response.status_code}")
    print(f"响应: {json.dumps(response.json(), indent=2, ensure_ascii=False)}")

    return response.json()


def test_get_result(shipment_id):
    """测试获取结果"""
    print(f"测试获取结果 {shipment_id}...")

    response = requests.get(f"{BASE_URL}/api/v1/shipments/{shipment_id}/result")
    print(f"状态码: {response.status_code}")
    print(f"响应: {json.dumps(response.json(), indent=2, ensure_ascii=False)}")

    return response.json()


if __name__ == "__main__":
    # 测试健康检查
    test_health_check()

    # 测试完整流程（需要准备测试文件）
    # upload_result = test_upload_files()
    # shipment_id = upload_result['data']['shipment_id']
    # task_id = upload_result['data']['task_id']

    # process_result = test_process_shipment(shipment_id)
    # task_id = process_result['data']['task_id']

    # import time
    # time.sleep(5)  # 等待处理完成

    # test_task_status(task_id)
    # test_get_result(shipment_id)
