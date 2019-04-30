package com.buildword;

//author herbert QQ 464884492
public class TagInfo {
	public String TagText; // ��${}
	public String TagValue;

	public TagInfo() {

	}

	public TagInfo(String tagText, String tagValue) {
		this.TagText = tagText;
		this.TagValue = tagValue;
	}

	@Override
	public boolean equals(Object x) {
		TagInfo i = (TagInfo) x;
		return i.TagText.equalsIgnoreCase(this.TagText);
	}
}
