package com.jp.test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 1.0v created by admin on 2021-1-26
 */
public class TestPattern {
  public static void main(String[] args) {
    String  pat= "\\{\\{( |#|@|\\*)?[\\w\\u4e00-\\u9fa5]+(\\.[\\w\\u4e00-\\u9fa5]+)*\\}\\}";
    // String  pat= "(\\{\\{)|(\\}\\})";

    Pattern p = Pattern.compile(pat);
    String msg = "{{bbbb}}aaa{{name}}bbbb{{value}}";
    System.out.println(p.matcher(msg).matches());
    Matcher m = p.matcher(msg);

    System.out.println(m.groupCount());
   /* StringBuffer sb = new StringBuffer();
    while (m.find()) {
      // System.out.println(m.group()+"---" + msg.substring(m.start()+2,m.end()-2));

      // msg = msg.replace(m.group(),"ccc");
      m.appendReplacement(sb, "dog");

      // System.out.println(m.group());
    }
    System.out.println(sb.toString());
    System.out.println(msg);*/
  }
}
