package com.dw.vsd2png;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

public class XWPFVisioData extends POIXMLDocumentPart {
    private static final int DEFAULT_MAX_VISIO_SIZE = 100_000_000;
    private static int MAX_VISIO_SIZE = DEFAULT_MAX_VISIO_SIZE;

    public static void setMaxVisioSize(int length) {
        MAX_VISIO_SIZE = length;
    }

    /**
     * @return the max image size allowed for XSSF pictures
     */
    public static int getMaxVisioSize() {
        return MAX_VISIO_SIZE;
    }

    private Long checksum;
    public XWPFVisioData() {
        super();
    }


    public XWPFVisioData(PackagePart part) {
        super(part);
    }

    @Override
    protected void onDocumentRead() throws IOException {
        super.onDocumentRead();
    }

    /**
     * Gets the picture data as a byte array.
     * <p>
     * Note, that this call might be expensive since all the picture data is copied into a temporary byte array.
     * You can grab the picture data directly from the underlying package part as follows:
     * <br>
     * <code>
     * InputStream is = getPackagePart().getInputStream();
     * </code>
     * </p>
     *
     * @return the Picture data.
     */
    public byte[] getData() {
        try (InputStream stream = getPackagePart().getInputStream()) {
            return IOUtils.toByteArrayWithMaxLength(stream, getMaxVisioSize());
        } catch (IOException e) {
            throw new POIXMLException(e);
        }
    }

    /**
     * Returns the file name of the image, eg image7.jpg . The original filename
     * isn't always available, but if it can be found it's likely to be in the
     * CTDrawing
     */
    public String getFileName() {
        String name = getPackagePart().getPartName().getName();
        return name.substring(name.lastIndexOf('/') + 1);
    }
    public Long getChecksum() {
        if (this.checksum == null) {
            try (InputStream is = getPackagePart().getInputStream()) {
                this.checksum = IOUtils.calculateChecksum(is);
            } catch (IOException e) {
                throw new POIXMLException(e);
            }
        }
        return this.checksum;
    }

    @Override
    public boolean equals(Object obj) {
        /*
         * In case two objects ARE equal, but its not the same instance, this
         * implementation will always run through the whole
         * byte-array-comparison before returning true. If this will turn into a
         * performance issue, two possible approaches are available:<br>
         * a) Use the checksum only and take the risk that two images might have
         * the same CRC32 sum, although they are not the same.<br>
         * b) Use a second (or third) checksum algorithm to minimise the chance
         * that two images have the same checksums but are not equal (e.g.
         * CRC32, MD5 and SHA-1 checksums, additionally compare the
         * data-byte-array lengths).
         */
        if (obj == this) {
            return true;
        }

        if (obj == null) {
            return false;
        }

        if (!(obj instanceof XWPFPictureData)) {
            return false;
        }

        XWPFPictureData picData = (XWPFPictureData) obj;
        PackagePart foreignPackagePart = picData.getPackagePart();
        PackagePart ownPackagePart = this.getPackagePart();

        if ((foreignPackagePart != null && ownPackagePart == null)
                || (foreignPackagePart == null && ownPackagePart != null)) {
            return false;
        }

        if (ownPackagePart != null) {
            OPCPackage foreignPackage = foreignPackagePart.getPackage();
            OPCPackage ownPackage = ownPackagePart.getPackage();

            if ((foreignPackage != null && ownPackage == null)
                    || (foreignPackage == null && ownPackage != null)) {
                return false;
            }
            if (ownPackage != null) {

                if (!ownPackage.equals(foreignPackage)) {
                    return false;
                }
            }
        }

        Long foreignChecksum = picData.getChecksum();
        Long localChecksum = getChecksum();

        if (localChecksum == null) {
            if (foreignChecksum != null) {
                return false;
            }
        } else {
            if (!(localChecksum.equals(foreignChecksum))) {
                return false;
            }
        }
        return Arrays.equals(this.getData(), picData.getData());
    }

    @Override
    public int hashCode() {
        Long checksum = getChecksum();
        return checksum == null ? super.hashCode() : checksum.hashCode();
    }

    /**
     * *PictureData objects store the actual content in the part directly without keeping a
     * copy like all others therefore we need to handle them differently.
     */
    @Override
    protected void prepareForCommit() {
        // do not clear the part here
    }
}
