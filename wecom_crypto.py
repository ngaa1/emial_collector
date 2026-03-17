import base64
import string
import random
import hashlib
import time
import struct
import json
from Crypto.Cipher import AES


class FormatException(Exception):
    pass


class PKCS7Encoder:
    """提供基于PKCS7算法的加解密接口"""
    
    block_size = 32
    
    def encode(self, text):
        """ 对需要加密的明文进行填充补位
        @param text: 需要进行填充补位操作的明文
        @return: 补齐明文字符串
        """
        text_length = len(text)
        amount_to_pad = self.block_size - (text_length % self.block_size)
        if amount_to_pad == 0:
            amount_to_pad = self.block_size
        pad = bytes([amount_to_pad])
        return text + pad * amount_to_pad
    
    def decode(self, decrypted):
        """删除解密后明文的补位字符
        @param decrypted: 解密后的明文
        @return: 删除补位字符后的明文
        """
        if len(decrypted) == 0:
            raise FormatException("invalid pkcs7 payload")
        pad = decrypted[-1]
        if not pad or pad < 1 or pad > self.block_size or pad > len(decrypted):
            raise FormatException("invalid pkcs7 padding")
        for i in range(1, pad + 1):
            if decrypted[-i] != pad:
                raise FormatException("invalid pkcs7 padding")
        return decrypted[:-pad]


class Prpcrypt:
    """提供接收和推送给企业微信消息的加解密接口"""
    
    def __init__(self, key):
        self.key = key
        self.mode = AES.MODE_CBC
    
    def get_random_str(self):
        """ 随机生成16位字符串
        @return: 16位字符串
        """
        rule = string.ascii_letters + string.digits
        str_list = random.sample(rule, 16)
        return "".join(str_list).encode('utf-8')
    
    def encrypt(self, text, receiveid):
        """对明文进行加密
        @param text: 需要加密的明文
        @return: 加密得到的字符串
        """
        if isinstance(text, str):
            text = text.encode('utf-8')
        if isinstance(receiveid, str):
            receiveid = receiveid.encode('utf-8')
        
        # 生成16字节随机数
        random16 = self.get_random_str()
        
        # 计算消息长度（4字节，大端序）
        msg_len = struct.pack(">I", len(text))
        
        # 拼接数据
        raw = random16 + msg_len + text + receiveid
        
        # PKCS7 填充
        pkcs7 = PKCS7Encoder()
        padded = pkcs7.encode(raw)
        
        # AES-256-CBC 加密
        iv = self.key[:16]
        cryptor = AES.new(self.key, self.mode, iv)
        ciphertext = cryptor.encrypt(padded)
        
        # Base64 编码
        return base64.b64encode(ciphertext).decode('utf-8')


class WeComCrypto:
    """企业微信消息加密类"""
    
    def __init__(self, token, encoding_aes_key, receiveid):
        try:
            # 确保 encoding_aes_key 以 "=" 结尾
            trimmed = encoding_aes_key.strip()
            if not trimmed:
                raise FormatException("encodingAESKey missing")
            with_padding = trimmed if trimmed.endswith("=") else trimmed + "="
            self.key = base64.b64decode(with_padding)
            if len(self.key) != 32:
                raise FormatException(f"invalid encodingAESKey (expected 32 bytes after base64 decode, got {len(self.key)})")
        except Exception as e:
            raise FormatException(f"EncodingAESKey 无效: {str(e)}")
        self.token = token
        self.receiveid = receiveid
    
    def compute_signature(self, token, timestamp, nonce, encrypt):
        """计算企业微信消息签名"""
        try:
            parts = [str(token or ""), str(timestamp or ""), str(nonce or ""), str(encrypt or "")]
            parts.sort()
            sha = hashlib.sha1()
            sha.update("" .join(parts).encode('utf-8'))
            return sha.hexdigest()
        except Exception as e:
            raise Exception(f"计算签名失败: {str(e)}")
    
    def encrypt_message(self, message, nonce=None, timestamp=None):
        """
        加密消息
        @param message: 消息内容（字典或字符串）
        @param nonce: 随机字符串
        @param timestamp: 时间戳
        @return: (加密后的数据, msg_signature, timestamp, nonce)
        """
        if isinstance(message, dict):
            message = json.dumps(message, ensure_ascii=False)
        
        if nonce is None:
            rule = string.ascii_letters + string.digits
            nonce = "".join(random.sample(rule, 16))
        
        if timestamp is None:
            timestamp = str(int(time.time()))
        
        pc = Prpcrypt(self.key)
        encrypt = pc.encrypt(message, self.receiveid)
        
        signature = self.compute_signature(self.token, timestamp, nonce, encrypt)
        
        return encrypt, signature, timestamp, nonce
