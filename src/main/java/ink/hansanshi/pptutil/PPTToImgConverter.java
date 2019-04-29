package ink.hansanshi.pptutil;


import net.coobird.thumbnailator.Thumbnails;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

/**
 * 这是一个工具类，可以将ppt或者pptx文件转换为图片
 *
 * 使用Apache POI将幻灯片转换为图片，但是直接转换的图片质量很低.
 * 对其进行放大转换后质量会有所改善，但是文件体积也会比较大。
 * 所以基本思路是：利用poi放大转换幻灯片--> 利用thumbnailator进行压缩并输出
 * @author ：hansanshi
 * @date ：Created in 2019/4/29
 */
public class PPTToImgConverter {

    /**支持四种图片格式，jpg,png,bmp,wbmp*/
    public static final String JPG = ".jpg";
    public static final String PNG = ".png";
    public static final String BMP = ".bmp";
    public static final String WBMP = ".wbmp";

    /**.ppt格式*/
    private static final int PPT_FORMAT = 0;
    /**.pptx格式*/
    private static final int PPTX_FORMAT = 1;
    /**其他文件格式*/
    private static final int OTHER_FORMAT = -1;

    /**输出图片的格式，默认为jpg*/
    private String imgFormat = JPG;
    /**临时图片放大倍数，默认为3*/
    private int scale = 3;
    /**压缩比例，如指定了宽度或者高度，则该值不会生效*/
    private Double ratio = 1.00;
    /**需要转换的文件列表*/
    private File[] files;
    /**输出图片分辨率的宽度*/
    private Integer width;
    /**输出图片分辨率的高度*/
    private Integer height;
    /**设定幻灯片转换的起始页码*/
    private  int start = 0;
    /**设定幻灯片转换的结束页码，该页会被转换，如果是负值，代表倒数第几页*/
    private  int end = -1;
    /**图片输出目录前缀*/
    private File outputDirPrefix;

    /**设置操作目录*/
    public static PPTToImgConverter setDir(String dirPrefix){
        return new PPTToImgConverter(dirPrefix);
    }
    /**选择单个文件或者文件夹*/
    public PPTToImgConverter loadFile(String filePath){
        return loadFile(new File(filePath));
    }
    /**选择单个文件或者文件夹*/
    public PPTToImgConverter loadFile(File file){
        return loadFiles(new File[]{file});
    }
    /**选择多个文件或者文件夹*/
    public PPTToImgConverter loadFiles(String[] filePaths){
        File[] files = new File[filePaths.length];
        for (int i = 0; i < filePaths.length; i++) {
            files[i] = new File(filePaths[i]);
        }
        return loadFiles(files);
    }
    /**选择多个文件或者文件夹*/
    public PPTToImgConverter loadFiles(File[] files){
        List<File> slideshowFiles = new ArrayList<>();
        for (File file:files){
            if (file.isDirectory()){
                slideshowFiles.addAll(getSlideShowFiles(file));
            }else{
                if (getFormat(file)>OTHER_FORMAT){
                    slideshowFiles.add(file);
                }
            }
        }
        this.files = slideshowFiles.toArray(new File[0]);
        return this;
    }


    /**
     * 设置临时图片放大倍数，默认为3
     * @param scale 临时图片放大倍数
     * @return this
     */
    public PPTToImgConverter setScale(int scale){
        if (scale<=0){
            throw new IllegalArgumentException("scale must be positive");
        }
        this.scale = scale;
        return this;
    }

    /**
     * 设置输出图片分辨率
     * @param width 分辨率宽度
     * @param height 分辨率高度
     * @return this
     */
    public PPTToImgConverter setSize(int width, int height){
        return this.setWidth(width).setHeight(height);
    }

    /**
     * 设置输出图片分辨率宽度
     * @param width 分辨率宽度
     * @return this
     */
    public PPTToImgConverter setWidth(int width){
        if (width<=0){
            throw new IllegalArgumentException("width must be positive");
        }
        this.width = width;
        return this;
    }

    /**
     * 设置输出图片分辨率高度
     * @param height 分辨率高度
     * @return this
     */
    public PPTToImgConverter setHeight(int height){
        if (height<=0){
            throw new IllegalArgumentException("height must be positive");
        }
        this.height = height;
        return this;
    }

    /**
     * 设置输出图片压缩比例，相对于原幻灯片分辨率
     * @param ratio 压缩比例
     * @return this
     */
    public PPTToImgConverter setRatio(double ratio){
        if (ratio <= Double.MIN_VALUE){
            throw new IllegalArgumentException("ratio must be positive");
        }
        this.ratio = ratio;
        return this;
    }

    /**
     * 设置转换起始页，从0开始
     * @param start 起始页
     * @return this
     */
    public PPTToImgConverter from(int start){
        this.start = start;
        return this;
    }

    /**
     * 设置转换结束页，该页也会被转换
     * @param end 结束页
     * @return this
     */
    public PPTToImgConverter to(int end){
        this.end = end;
        return this;
    }

    public PPTToImgConverter setImgFormat(String format){
        if (JPG.equals(format)
        ||PNG.equals(format)
        ||BMP.equals(format)
        ||WBMP.equals(format)){
            this.imgFormat = format;
            return this;
        }else{
            throw new IllegalArgumentException("Image format not supported");
        }
    }

    /**开始转换*/
    public void convert()  {
        if (start*end > 0 && end<start){
            throw new IllegalArgumentException("start position is behind end!");
        }
        this.convertFiles();
    }

    /**只转换第一页*/
    public void convertFirstPage(){
        this.convertCertainPage(0);
    }

    /**只转换最后一页*/
    public void convertLastPage(){
        this.convertCertainPage(-1);
    }

    /**只转换特定一页*/
    public void convertCertainPage(int pageNum){
        this.from(pageNum)
                .to(pageNum)
                .convert();
    }

    /**转换所有页*/
    public void convertAllPages(){
        this.from(0).to(-1)
                .convert();
    }

    /**
     * 调用构造方法，设置输出目录前缀
     * @param dirPrefix 输出目录前缀
     */
    private PPTToImgConverter(String dirPrefix){
        this.outputDirPrefix = new File(dirPrefix,"output");
    }

    /**
     * 转换多个文件并输出转换结果
     */
    private void convertFiles()  {
        for (File file:files) {
            File[] outputImgs;
            try {
                outputImgs = convertSlideshow(file);
            } catch (Exception e) {
                System.out.println(file.getAbsolutePath()+" can't be converted: "+e.getMessage());
                continue;
            }
            if (outputImgs.length == 0){
                System.out.println(file.getAbsolutePath()+" the result of conversion is null");
                continue;
            }
            System.out.println(file.getAbsolutePath()+" has been converted in: "+outputImgs[0].getParent());
        }
    }

    /**
     *转换幻灯片为图片
     * @param file 幻灯片文件
     * @return File[] 转换的图片文件路径列表
     * @throws IOException 多出可能抛出
     */
    private File[] convertSlideshow(File file) throws IOException {
        //将文件加载为幻灯片
        SlideShow slideShow= getSlideShow(file);

        //根据设置生成图片
        File outputDir = new File(outputDirPrefix,file.getName());
        File[] outputImgs = this.generateImg(outputDir,slideShow);

        return outputImgs;
    }

    /**
     * 获取幻灯片，并根据幻灯片信息设置输出图片的分辨率
     * @param file 幻灯片文件
     * @return SlideShow 幻灯片
     * @throws IOException
     */
    private SlideShow getSlideShow(File file) throws IOException {

        //根据不同文件格式进行初始化
        FileInputStream inputStream = new FileInputStream(file);
        SlideShow slideShow;
        int fileFormat = getFormat(file);
        if (fileFormat == PPTToImgConverter.PPTX_FORMAT){
            slideShow = new XMLSlideShow(inputStream);
        }else if (fileFormat == PPTToImgConverter.PPT_FORMAT){
            slideShow = new HSLFSlideShow(inputStream);
        }else{
            inputStream.close();
            throw new IOException("Not SlideShow");
        }
        inputStream.close();

        return slideShow;
    }

    /**
     * 确定需要的页码并生成对应图片
     * @param outputDir 用于存放图片的目录
     * @param slideShow 幻灯片
     * @return File[] 生成的图片文件列表
     * @throws IOException 生成图片时可能抛出
     */
    private File[] generateImg(File outputDir,
                               SlideShow slideShow) throws IOException {
        //创建输出目录
        while(outputDir.exists()){
            outputDir = new File(outputDir.getParent(),
                    UUID.randomUUID().toString());
        }
        outputDir.mkdirs();


        //生成每个文件实际的起始和结束位置
        //如果start和end的设置值为负的话，表示是是倒数第几页，例-1表示最后一页
        List<Slide> slides = slideShow.getSlides();
        int realStart = start<0?start+slides.size():start;
        int realEnd = end<0?end+slides.size():end;
        //起始点不在范围内或者结束位置不合理
        if (realStart<0 || realEnd<0 || realStart>= slides.size() || realStart > realEnd){
            return new File[0];
        }
        //只要起始点处于有效范围即可，终点值如超出范围则置为最后一页
        if (realEnd >= slides.size()){
            realEnd = slides.size()-1 ;
        }

        //根据幻灯片的分辨率和预先的设置以生成实际输出图片的分辨率
        //一旦设置width和height，则ratio的设置无效
        Dimension pageSize = slideShow.getPageSize();
        int realWidth;
        int realHeight;
        if (width == null){
            realHeight = height==null?(int)(pageSize.height*ratio):height;
            realWidth = height==null?(int)(pageSize.width*ratio):height*pageSize.width/pageSize.height;
        }else{
            realWidth = width;
            realHeight = height==null? width*pageSize.height/pageSize.width:height;
        }

        //放大特定倍数倍以提高图片清晰度，图片文件体积也会随之增大
        BufferedImage bufferedImg = new BufferedImage(pageSize.width* scale,
                pageSize.height* scale,BufferedImage.TYPE_INT_RGB);
        File[] outputImgs = new File[realEnd-realStart+1];
        //开始逐页生成图片
        Graphics2D graphics = bufferedImg.createGraphics();
        graphics.scale(scale, scale);
        graphics.setPaint(Color.white);
        for (int i = realStart; i <= realEnd; i++) {
            //即用刷填充背景为白色，否则如果ppt背景是透明的话，会有上次绘图留下的背景
            graphics.fill(new Rectangle2D.Float(0, 0,
                    pageSize.width * scale, pageSize.height * scale));
            slides.get(i).draw(graphics);
            //压缩并保存为文件
            File outputImg = new File(outputDir,i+imgFormat);
            Thumbnails.of(bufferedImg)
                    .size(realWidth,realHeight)
                    .toFile(outputImg);
            outputImgs[i-realStart]=outputImg;
        }
        return outputImgs;
    }

    /**
     * 检测文件的格式
     * @param file 待检测格式的文件
     * @return int 文件的格式，分三种情况: OTHER_FORMAT, -1 ;PPT_FORMAT, 0; PPTX_FORMAT, 1
     */
    private int getFormat(File file) {
        int format = OTHER_FORMAT;
        String filename = file.getName();
        if (filename.contains(".") ) {
            String suffix = filename.substring(filename.lastIndexOf(".")).toLowerCase();
            if (".ppt".equals(suffix)) {
                format = PPT_FORMAT;
            }else if (".pptx".equals(suffix)) {
                format = PPTX_FORMAT;
            }
        }
        return format;
    }

    /**
     * 递归获取目录及其子目录的幻灯片文件
     * @param dir 目录
     * @return 幻灯片文件列表 不返回Null
     */
    private List<File> getSlideShowFiles(File dir) {
        List<File> files=new ArrayList<>();
        if (dir.isDirectory()){
            File[] subFiles = dir.listFiles();
            if (subFiles == null){
                return files;
            }
            for (File subFile: subFiles) {
                if (subFile.isDirectory()){
                    //获取子目录的幻灯片文件
                    files.addAll(getSlideShowFiles(subFile));
                }else{
                    //检测文件格式，如果是幻灯片，则加入列表中
                    int result = getFormat(subFile);
                    if (result > OTHER_FORMAT){
                        files.add(subFile);
                    }
                }
            }
        }else{
            int result = getFormat(dir);
            if (result > OTHER_FORMAT){
                files.add(dir);
            }
        }
        return files;
    }

}
