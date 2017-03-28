<?php
/**
 * 对于数据量很多的网站，运营人员都习惯使用Excel程序进行数据的修改和维护，
 * 他们希望网站可以将他们修改后的Excel文件导入到我们的网站上，这样可以提高工作效率。
 * 但这就会产生一个问题，因为通常Excel文件是二进制文件，程序代码都不可能读取到里面
 * 的数据，所以就产生了一个折中的办法，就是使用Excel可以编辑，并且程序代码也可以
 * 读取的文件类型（xml，csv等）。
 *
 * 运营人员需要先将Excel保存为 xml或csv格式，然后再导入到网站数据库中。
 * @author panke
 */

/**
 * 定义抽象类 Excel
 */
abstract class Excel {

    /**
     * 初始化方法
     */
    public static function init($className) {
        //判断类是否存在
        if (!class_exists($className)) {
            $this->error('请确定文件类型是否正确');
        }

        $instance = new $className();
        if($instance instanceof self){
            return $instance;
        }
        return null;
    }

    /**
     * 输出错误消息
     */
    protected function error($msg) {
        echo $msg;
        exit;
    }
}

/**
 * 定义抽象Excel导入类
 */
abstract class ExcelImport extends Excel {
    //数据资源
    protected $source;
    //数据行数
    public $rowsNum;

    /**
     * 初始化方法
     */
    public final static function init($type) {
         //类名
        $className = ucfirst(strtolower($type)).'Import';
        return parent::init($className);
    }

    /**
     * 加载资源文件
     */
    abstract function loadFile($sourceFile);

    /**
     * 加载资源字符串
     */
    abstract function loadString($sourceString);

    /**
     * 获取一行数据
     */
    abstract function getRow($num);

    /**
     * 获取一列数据
     */
    //abstract function getColumn();

    /**
     * 获取单元格数据
     */
    //abstract function getCell();

}

/**
 * 定义抽象Excel导出类
 */
abstract class ExcelExport extends Excel {
    //数据
    protected $datas;

    /**
     * 初始化方法
     */
    public final static function init($type) {
         //类名
        $className = ucfirst(strtolower($type)).'Export';
        return parent::init($className);
    }

     /**
     * 设置文档标题
     */
    //abstract function setTitle();

    /**
     * 设置一行数据
     */
    abstract function setRow($row);

    /**
     * 显示数据
     */
    abstract function show();
}

/**
 * 定义Xml导入类，并继承ExcelImport抽象类
 */
class XmlImport extends ExcelImport {
    /**
     * xml的命名空间
     */
    protected $nameSpace;

    /**
     * 实现加载资源文件方法
     */
    public function loadFile($sourceFile) {
        if (!file_exists($sourceFile)) {
            $this->error('文件不存在');
        }

        $sourceString = file_get_contents($sourceFile);
        $this->loadString($sourceString);
    }

    /**
     * 实现加载资源字符串方法
     */
    public function loadString($sourceString) {
        $xml = simplexml_load_string($sourceString,'SimpleXMLElement',LIBXML_NOCDATA);
        
        $worksheet = $xml->Worksheet;
        $table = $worksheet->Table;
        $rows = $table->Row;

        $this->source = $rows;

        $this->rowsNum = count($rows);
        $nameSpaces = $xml->getDocNamespaces();
        $this->nameSpace = $nameSpaces['ss'];

    }

    /**
     * 实现获取一行数据方法
     */
    public function getRow($num) {

        if ($num > $this->rowsNum || $num <= 0) {
            $this->error('该行('.$num.')数据不存在');
        }

        return $this->rowToArray($this->source[$num - 1]);
    }

    /**
     * 将一行数据变成数组
     */
    protected function rowToArray($row) {
        $cells = $row->Cell;

        $tmp = array();
        foreach ($cells as $cell) {
            $cell->registerXPathNamespace('prefix', $this->nameSpace);

            /**
             * 以下代码为了处理某单元格无数据，且没有该单元格的 <Cell>标签时的情况
             */
            $attribute = $cell->xpath("attribute::prefix:Index");
            if (!empty($attribute)) {
                $index = $attribute[0]->Index;
                $tmp = array_pad($tmp, $index-1, '');
            }

            /**
             * 处理Data标签带前缀的情况，且Data标签内容包含样式标签的情况
             */
            $data = $cell->xpath("child::prefix:Data");
            $tmp[] = isset($data[0]) ? strip_tags($data[0]->asXML()) : '';
        }

        return $tmp;
    }
}

/**
 * XML数据导出类
 */
class XmlExport extends ExcelExport {
    /**
     * 设置一行数据
     */
    public function setRow($row) {
        if (!is_array($row)) {
            $this->error('请传入数组数据');
        }

        $line = '<Row>';
        foreach ($row as $cell) {
            $line .= '<Cell><Data ss:Type="String">'.$cell."</Data></Cell>\n";
        }
        $line .= "</Row>\n";
        $this->datas .= $line;
    }

    /**
     * 显示数据
     */
    public function show() {
        //前缀
        $prefix = '<?xml version="1.0" encoding="UTF-8"?>
            <?mso-application progid="Excel.Sheet"?>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:x="urn:schemas-microsoft-com:office:excel"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:html="http://www.w3.org/TR/REC-html40">
            <Worksheet ss:Name="Table1">
            <Table>
        ';

        //后缀
        $suffix = '
            </Table></Worksheet></Workbook>
        ';

        return $prefix.$this->datas.$suffix;
    }
}

/**
 * 定义Csv类，并继承Excel抽象类
 */
//class Csv extends Excel {
//
//}


