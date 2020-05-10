package com.qxtx.idea.ideaexcel.poi.bean;

/**
 * Created in 2020/4/26 17:09
 *
 * @author QXTX-WORK
 * <p>
 * Description 表格中行内容实体类
 */
public class RowBean {

    /** 姓名 */
    private String name;

    /** 证件号码 */
    private String id;

    /** 户籍类别 */
    private String censusType;

    /** 居住地址 */
    private String address;

    public RowBean() {}

    public RowBean(String name, String id, String censusType, String address) {
        this.name = name;
        this.id = id;
        this.censusType = censusType;
        this.address = address;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getCensusType() {
        return censusType;
    }

    public void setCensusType(String censusType) {
        this.censusType = censusType;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    @Override
    public String toString() {
        return "RowBean{" +
                "name='" + name + '\'' +
                ", id='" + id + '\'' +
                ", censusType='" + censusType + '\'' +
                ", address='" + address + '\'' +
                '}';
    }
}
