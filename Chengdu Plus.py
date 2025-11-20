import requests
import time
import pandas as pd
from openpyxl import Workbook
import sys  # 引入 sys 模块，用于处理编码

# 确保控制台可以正确打印中文
# reload(sys)
# sys.setdefaultencoding('utf-8')


# --- 配置区 ---
API_KEY = "AIzaSyAI45JaRlujtx8PgJG9wh5O5opdOJqX7P4"  # ← 在这里填入你的 API Key
CHANNEL_NAME = "Chengdu Plus"
OUTPUT_EXCEL = "chengdu_plus_videos_and_comments.xlsx"


# ----------------

# -----------------------------------------------------------
# 1. 获取频道的“上传”播放列表 ID (Uploads Playlist ID)
#    通过 Search 找到 Channel ID, 再通过 Channels 找到 Uploads Playlist ID
# -----------------------------------------------------------
def get_uploads_playlist_id(channel_name):
    # 第一步：搜索频道获取 Channel ID
    search_url = "https://www.googleapis.com/youtube/v3/search"
    search_params = {
        "part": "snippet",
        "q": channel_name,
        "type": "channel",
        "key": API_KEY,
        "maxResults": 1
    }
    try:
        r = requests.get(search_url, params=search_params).json()

        if "items" not in r or len(r["items"]) == 0:
            print(f"API 响应：{r}")
            raise Exception("未找到频道，请检查名称是否正确。")

        channel_id = r["items"][0]["snippet"]["channelId"]
        print(f"找到频道 ID: {channel_id}")

        # 第二步：获取该频道的 Uploads 播放列表 ID
        channel_url = "https://www.googleapis.com/youtube/v3/channels"
        channel_params = {
            "part": "contentDetails",
            "id": channel_id,
            "key": API_KEY
        }
        r_channel = requests.get(channel_url, params=channel_params).json()

        # 提取 uploads ID
        uploads_id = r_channel["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]
        return uploads_id

    except Exception as e:
        print(f"获取播放列表 ID 失败: {e}")
        return None


# -----------------------------------------------------------
# 2. 通过播放列表获取所有视频 ID (稳定且省配额)
#    替换了原有的 get_all_video_ids 函数
# -----------------------------------------------------------
def get_all_video_ids_from_playlist(playlist_id):
    url = "https://www.googleapis.com/youtube/v3/playlistItems"
    all_video_ids = []
    next_page = None

    print("开始获取视频列表 ID，这可能需要一些时间...")

    while True:
        params = {
            "part": "contentDetails",  # 只需要 video ID
            "playlistId": playlist_id,
            "maxResults": 50,
            "key": API_KEY
        }

        if next_page:
            params["pageToken"] = next_page

        try:
            resp = requests.get(url, params=params).json()
            # print(f"获取视频 ID 的 API 响应：{resp}") # 调试用

            if "items" not in resp:
                print(f"API 响应错误: {resp}")
                break

            for item in resp["items"]:
                # 通过 contentDetails 获取 videoId
                all_video_ids.append(item["contentDetails"]["videoId"])

            # 打印进度
            print(f"已获取 {len(all_video_ids)} 个视频 ID...")

            next_page = resp.get("nextPageToken")
            if not next_page:
                break

            time.sleep(0.15)  # 稍微限速，避免 API 超限

        except Exception as e:
            print(f"获取视频 ID 失败: {e}")
            break

    return all_video_ids


# -----------------------------------------------------------
# 3. 获取每个视频的所有可用字段（视频详情）
#    此函数与您原有的基本相同
# -----------------------------------------------------------
def get_video_details(video_ids):
    all_data = []
    total_videos = len(video_ids)

    for i in range(0, total_videos, 50):  # videos.list 一次最多 50 个
        batch = video_ids[i:i + 50]
        url = "https://www.googleapis.com/youtube/v3/videos"
        params = {
            "part": "snippet,statistics,contentDetails,topicDetails,status",
            "id": ",".join(batch),
            "key": API_KEY
        }

        print(f"正在获取第 {i + 1} 到 {min(i + 50, total_videos)} 个视频的详情...")

        try:
            resp = requests.get(url, params=params).json()
            # print(f"获取视频详情的 API 响应：{resp}") # 调试用

            if "items" not in resp:
                print(f"API 响应错误: {resp}")
                # 检查是否是 API Key 或配额问题
                if resp.get("error"):
                    print(f"错误详情: {resp.get('error')}")
                break

            for item in resp["items"]:
                snippet = item.get("snippet", {})
                stats = item.get("statistics", {})
                content = item.get("contentDetails", {})

                all_data.append({
                    "videoId": item.get("id"),
                    "title": snippet.get("title"),
                    "description": snippet.get("description"),
                    "publishedAt": snippet.get("publishedAt"),
                    "tags": ",".join(snippet.get("tags", [])) if "tags" in snippet else "",
                    "categoryId": snippet.get("categoryId"),

                    # statistics
                    "viewCount": stats.get("viewCount"),
                    "likeCount": stats.get("likeCount"),
                    "commentCount": stats.get("commentCount"),

                    # content details
                    "duration": content.get("duration"),
                    "definition": content.get("definition"),
                    "caption": content.get("caption"),

                    # status
                    "privacyStatus": item.get("status", {}).get("privacyStatus", ""),

                    # topicDetails
                    "topicCategories": ",".join(item.get("topicDetails", {}).get("topicCategories", []))
                })

            time.sleep(0.25)  # 限速

        except Exception as e:
            print(f"获取视频详情失败: {e}")
            break

    return all_data


# -----------------------------------------------------------
# 4. 获取视频的所有评论（带翻页）
#    此函数与您原有的基本相同
# -----------------------------------------------------------
def get_video_comments(video_id):
    comments = []
    next_page = None
    page_count = 0  # 追踪页数

    while True:
        url = "https://www.googleapis.com/youtube/v3/commentThreads"
        params = {
            "part": "snippet",
            "videoId": video_id,
            "maxResults": 100,  # 每页最大评论数
            "textFormat": "plainText",
            "key": API_KEY
        }

        if next_page:
            params["pageToken"] = next_page

        try:
            resp = requests.get(url, params=params).json()
            # print(f"获取评论的 API 响应：{resp}") # 调试用

            # 检查是否有错误信息，例如评论被禁用
            if resp.get("error"):
                # 如果错误是 "commentsDisabled" (403), 那么就跳过
                if resp["error"]["code"] == 403:
                    # print(f"视频 {video_id} 评论已禁用或不可用。")
                    return comments
                else:
                    print(f"获取评论时发生 API 错误: {resp['error']}")
                    break

            for item in resp.get("items", []):
                comment_data = item["snippet"]["topLevelComment"]["snippet"]
                comments.append({
                    "videoId": video_id,
                    "commentId": item["id"],
                    "author": comment_data["authorDisplayName"],
                    "publishedAt": comment_data["publishedAt"],
                    "text": comment_data["textDisplay"],
                    "likeCount": comment_data.get("likeCount", 0)
                })

            page_count += 1
            # print(f"视频 {video_id} 评论已获取 {page_count} 页 ({len(comments)} 条)...")

            next_page = resp.get("nextPageToken")
            if not next_page:
                break

            time.sleep(0.25)  # 限速

        except Exception as e:
            print(f"获取评论失败: {e}")
            break

    return comments


# -----------------------------------------------------------
# 5. 保存为 Excel 文件（视频数据 & 评论数据）
#    此函数与您原有的相同
# -----------------------------------------------------------
def save_to_excel(video_data, comment_data, filename):
    # 写入视频数据
    video_df = pd.DataFrame(video_data)

    # 写入评论数据
    comment_df = pd.DataFrame(comment_data)

    # 保存 Excel 文件
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        video_df.to_excel(writer, sheet_name="videos", index=False)
        comment_df.to_excel(writer, sheet_name="comments", index=False)

    print(f"已成功保存到文件：{filename}")


# -----------------------------------------------------------
# 主流程
# -----------------------------------------------------------
if __name__ == "__main__":
    print("正在查找频道及上传列表...")
    # 1. 获取上传列表 ID (替换了原有的 get_channel_id)
    uploads_playlist_id = get_uploads_playlist_id(CHANNEL_NAME)

    if not uploads_playlist_id:
        print("未能获取播放列表 ID，程序退出。")
        sys.exit(1)

    print("\n====================================")
    print("正在获取所有视频 ID...")
    # 2. 使用新的函数获取视频 ID (使用 playlistItems)
    video_ids = get_all_video_ids_from_playlist(uploads_playlist_id)
    print(f"共获取到 {len(video_ids)} 个视频。")
    print("====================================")

    if not video_ids:
        print("未获取到任何视频 ID，程序退出。")
        sys.exit(1)

    print("正在获取视频详细信息...")
    video_data = get_video_details(video_ids)
    print(f"已获取 {len(video_data)} 个视频的详情数据。")

    print("\n====================================")
    print("正在获取视频评论 (注意：此步骤最耗费时间，也最消耗 API 配额)...")
    comment_data = []

    # 建议先测试少量，例如只抓取前100个视频的评论: video_ids_to_process = video_ids[:100]
    video_ids_to_process = video_ids  # 处理全部视频

    for i, video_id in enumerate(video_ids_to_process):
        print(f"[{i + 1}/{len(video_ids_to_process)}] 正在处理视频 {video_id} 的评论...")
        comments = get_video_comments(video_id)
        comment_data.extend(comments)

    print("====================================")
    print(f"所有视频共获取到 {len(comment_data)} 条评论。")

    print("正在保存为 Excel 文件...")
    save_to_excel(video_data, comment_data, OUTPUT_EXCEL)

    print("\n全部完成！🎉")