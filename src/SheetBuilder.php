<?php
/**
 * Created by PhpStorm.
 * User: wim
 * Date: 2018/10/24
 * Time: 17:39
 */
namespace wim\export;

abstract class SheetBuilder
{
    /**
     * 设置表名称
     * @return mixed
     */
    public abstract function buildTitle();

    /**
     * 设置表头
     * @return mixed
     */
    public abstract function buildHeaders();

    /**
     * 设置表格内容
     */
    public abstract function buildContent();

    /**
     * 获取表格
     */
    public abstract function getResult();

}