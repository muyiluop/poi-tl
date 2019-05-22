package com.deepoove.poi.tl;

import java.util.regex.Pattern;

import org.junit.Assert;
import org.junit.Test;

import com.deepoove.poi.config.Configure;

public class RegTest {

    final String TAG_REGEX = Configure.newBuilder().build().getGramerPrefix();
    final String EL_REGEX = "^[^\\.]+(\\.[^\\\\.]+)*$";

    @Test
    public void testELReg() {
        Pattern pattern = Pattern.compile(EL_REGEX);
        testReg(pattern);
    }

    @Test
    public void testTagReg() {
        Pattern pattern = Pattern.compile(TAG_REGEX);
        testReg(pattern);
        Assert.assertFalse(pattern.matcher("abc-123").matches());
    }

    public void testReg(Pattern pattern) {
        Assert.assertTrue(pattern.matcher("123").matches());
        Assert.assertTrue(pattern.matcher("ABC").matches());
        Assert.assertTrue(pattern.matcher("abc123").matches());
        Assert.assertTrue(pattern.matcher("_123abc").matches());
        Assert.assertTrue(pattern.matcher("abc_123").matches());
        Assert.assertTrue(pattern.matcher("好123").matches());
        Assert.assertTrue(pattern.matcher("123好_好abc").matches());
        Assert.assertTrue(pattern.matcher("abc.123").matches());

        Assert.assertTrue(pattern.matcher("abc.123").matches());

        Assert.assertTrue(pattern.matcher("abc.123.123").matches());
        Assert.assertTrue(pattern.matcher("abc.好.123").matches());
        Assert.assertTrue(pattern.matcher("abc.好123").matches());
        Assert.assertTrue(pattern.matcher("好.123").matches());
        Assert.assertTrue(pattern.matcher("好.123.好").matches());

        Assert.assertFalse(pattern.matcher("好..123").matches());
        Assert.assertFalse(pattern.matcher("abc..123").matches());
        Assert.assertFalse(pattern.matcher("abc23.").matches());
        Assert.assertFalse(pattern.matcher("好123.").matches());
        Assert.assertFalse(pattern.matcher(".好123").matches());
    }

}
