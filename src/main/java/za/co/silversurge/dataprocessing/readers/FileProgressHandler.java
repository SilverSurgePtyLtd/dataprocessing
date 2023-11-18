package za.co.silversurge.dataprocessing.readers;

/**
 * Used to track the progress of an Excel file
 */
public interface FileProgressHandler {

    void onUpdate(long index, long max);

}
