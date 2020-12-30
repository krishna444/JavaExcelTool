package com.kpaudel;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * @author Krishna Paudel<br>
 * Date 20.10.20 11:33
 */
public class MD5Generator {

   public static void main(String... args) throws NoSuchAlgorithmException {
      MD5Generator generator = new MD5Generator();
      String[][] credentials = {{"Title2", "XXXX", "XXXX"},
              {"Title1", "XXXX", "XXXX"},
              {"Title3", "XXXX", "XXXXXX"},
              {"Title4", "XXXX", "XXXX"}
      };
      //String user="phastssodaprods";
      //String password="j2TiPGXsvRsNNFEk";
      for (String[] credential : credentials) {
         String hashedString = generator.generateMD5(String.format("%s:ApplicationRealm:%s", credential[1], credential[2]));
         System.out.println(credential[0] + " " + hashedString);
      }

   }

   public String generateMD5(String credential) throws NoSuchAlgorithmException {
      MessageDigest messageDigest = MessageDigest.getInstance("MD5");
      messageDigest.update(credential.getBytes(StandardCharsets.UTF_8));
      byte[] digest = messageDigest.digest();
      //String md5HashedText = new String(Base64.getEncoder().encode(digest));
      StringBuffer sb = new StringBuffer();
      for (int i = 0; i < digest.length; ++i) {
         sb.append(Integer.toHexString((digest[i] & 0xFF) | 0x100).substring(1, 3));
      }
      return sb.toString();
   }


}
