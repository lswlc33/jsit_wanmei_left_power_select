import requests
import threading
from datetime import datetime
from openpyxl import Workbook

# 速度设置（线程数）128大概半分钟
runSpeed = 128

def getLeftPower(roomNum):
    # 获取剩余电量，返回string
    url = "https://cloudpaygateway.59wanmei.com:8087/paygateway/smallpaygateway/trade"
    roomNum = str(roomNum)
    payload = {
        "method": "samllProgramGetRoomState",
        "bizcontent": '{"payproid":273,"schoolcode":"609","roomverify":"'
        + f"1-{8+int(roomNum[0])}--{roomNum[1]}-{roomNum}"
        + '","businesstype":2}',
    }

    response = requests.post(url, json=payload)

    if response.status_code == 200:
        data = response.json()
        power_left = data["businessData"]["quantity"]
        return power_left
    else:
        return "-99999"


def getRoomsList(buildno, floor):
    # 获取指定楼号指定楼层的所有宿舍号
    url = "https://cloudpaygateway.59wanmei.com:8087/paygateway/smallpaygateway/trade"

    payload = {
        "method": "samllProgramGetRoom",
        "bizcontent": '{"schoolno":"609","optype":"4","payproid":273,"areaid":"1","buildid":"'
        + f"{8+int(buildno)}"
        + '","unitid":"1","levelid":"'
        + str(floor)
        + '","businesstype":2}',
    }

    response = requests.post(url, json=payload)

    data = response.json()["businessData"]
    rooms = []
    for room in data:
        rooms.append(room["name"])
    return rooms


def proceRoomData(room):
    # log并添加数据到表格
    try:
        power = float(getLeftPower(room))
        if power < -500:
            status = "传说"
        elif power < 0:
            status = "死了"
        elif power < 20:
            status = "危"
        else:
            status = "正常"

        with lock:
            sheet.append([int(room[0]), int(room[1]), int(room), float(power), status])
            print(f"宿舍号: {room} \t 剩余电量: {power}\t 状况: {status}")
    except:
        return


def process_rooms(buildno, floor):
    rooms = getRoomsList(buildno, floor)
    for room in rooms:
        proceRoomData(room)


if __name__ == "__main__":
    fileName = datetime.now().strftime("%Y%m%d_%H_%M")
    header = ["楼号", "楼层", "宿舍号", "剩余电量", "状况"]
    new_excel = Workbook()
    sheet = new_excel.active
    buildnos = [1, 2, 3, 4, 5, 6, 7, 8]
    floors = [1, 2, 3, 4, 5, 6, 7]
    lock = threading.Lock()
    threads = []

    # 添加表头
    sheet.append(header)

    # 主循环
    for buildno in buildnos:
        for floor in floors:
            thread = threading.Thread(target=process_rooms, args=(buildno, floor))
            threads.append(thread)
            thread.start()

            while threading.active_count() > runSpeed:
                pass

    for thread in threads:
        thread.join()

    # 保持表格
    new_excel.save(f"{fileName}.xlsx")
    print(f"\n\n\n\n表格文件已生成 {fileName}.xlsx")
