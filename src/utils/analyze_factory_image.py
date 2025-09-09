from openai import OpenAI, APIError
import os
import base64

def encode_image(image_path):
    """
    将图像文件编码为Base64格式。
    
    参数:
    image_path (str): 图像文件的路径。
    
    返回:
    str: Base64编码的图像字符串。
    """
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")
    except FileNotFoundError:
        raise ValueError(f"图像文件 {image_path} 未找到。")
    except Exception as e:
        raise RuntimeError(f"读取图像文件 {image_path} 时发生错误: {e}")

def analyze_factory_image(image_path):
    """
    分析工厂产品宣传图，推测工厂类型、产品品类和具体产品。
    
    参数:
    image_path (str): 图像文件的路径。
    
    返回:
    dict: 分析结果。
    """
    base64_image = encode_image(image_path)
    
    client = OpenAI(
        api_key=os.getenv('DASHSCOPE_API_KEY1'),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )
    
    try:
        completion = client.chat.completions.create(
            model="qwen-vl-max-latest",
            messages=[
                {
                    "role": "system",
                    "content": [{"type": "text", "text": "You are a helpful assistant."}]
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                        },
                        {"type": "text", "text": "这是一张工厂产品宣传图，先推测这是什么类型的工厂，做什么产品？列举一些工厂可能做的品类和具体产品，最后具体描述图片中的具体实体产品。要求简明扼要，不超过100字"},
                    ],
                }
            ],
        )
        return completion.choices[0].message.content
    except APIError as e:
        raise RuntimeError(f"API调用失败: {e}")
    except Exception as e:
        raise RuntimeError(f"分析图像时发生错误: {e}")

# 示例调用
if __name__ == "__main__":
    image_path = "D:/文档/桌面/test-imgs/DM_20250310144113_001.png"
    try:
        result = analyze_factory_image(image_path)
        print(result)
    except Exception as e:
        print(f"错误: {e}")