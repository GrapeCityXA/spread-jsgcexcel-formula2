package com.grapecity.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;

public class GZip {
	public static String compress(String str) {
        if (str.length() <= 0) {
            return str;
        }
        try {
            ByteArrayOutputStream bos = null;
            GZIPOutputStream os = null; //使用默认缓冲区大小创建新的输出流
            byte[] bs = null;
            try {
                bos = new ByteArrayOutputStream();
                os = new GZIPOutputStream(bos);
                os.write(str.getBytes()); //写入输出流
                os.close();
                bos.close();
                bs = bos.toByteArray();
                return new String(bs, "ISO-8859-1");
            } finally {
                bs = null;
                bos = null;
                os = null;
            }
        } catch (Exception ex) {
            return str;
        }
    }

    public static String uncompress(String str) {
        if (str.length() <= 0) {
            return str;
        }
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ByteArrayInputStream in = new ByteArrayInputStream(str.getBytes("ISO-8859-1"));
            GZIPInputStream ungzip = new GZIPInputStream(in);
            byte[] buffer = new byte[256];
            int n;
            while ((n = ungzip.read(buffer)) >= 0) {
                out.write(buffer, 0, n);
            }
            return new String(out.toByteArray(), "UTF-8");
        } catch (Exception e) {

        }
        return str;
    }
}
