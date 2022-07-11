from PCANBasic import *
from ManualRead import *

m_objPCANBasic = PCANBasic()


def GetDataString(data, msgtype):
    """
        Gets the data of a CAN message as a string

        Parameters:
            data = Array of bytes containing the data to parse
            msgtype = Type flags of the message the data belong

        Returns:
            A string with hexadecimal formatted data bytes of a CAN message
        """
    if (msgtype & PCAN_MESSAGE_RTR.value) == PCAN_MESSAGE_RTR.value:
        return "Remote Request"
    else:
        strTemp = b""
        for x in data:
            strTemp += b'%.2X ' % x
        return str(strTemp).replace("'", "", 2).replace("b", "", 1)
        # return strTemp


def GetTimeString(time):
    """
        Gets the string representation of the timestamp of a CAN message, in milliseconds

        Parameters:
            time = Timestamp in microseconds

        Returns:
            String representing the timestamp in milliseconds
        """
    fTime = time / 1000
    return '%.1f' % fTime


def ProcessMessageCan(msg, itstimestamp):
    """
         Processes a received CAN message
         
         Parameters:
             msg = The received PCAN-Basic CAN message
             itstimestamp = Timestamp of the message as TPCANTimestamp structure
         """
    microsTimeStamp = itstimestamp.micros + 1000 * itstimestamp.millis + 0x100000000 * 1000 * itstimestamp.millis_overflow

    # print("Type: " + GetTypeString(msg.MSGTYPE))
    # print("ID: " + GetIdString(msg.ID, msg.MSGTYPE))
    # print("Length: " + str(msg.LEN))
    read_time = GetTimeString(microsTimeStamp)
    read_data = GetDataString(msg.DATA, msg.MSGTYPE)

    return read_time, read_data


'''
def ReadMessage(self):
    """
        Function for reading CAN messages on normal CAN devices

        Returns:
            A TPCANStatus error code
        """
    ## We execute the "Read" function of the PCANBasic
    stsResult = self.m_objPCANBasic.Read(self.PcanHandle)

    if stsResult[0] == PCAN_ERROR_OK:
        ## We show the received message
        self.ProcessMessageCan(stsResult[1], stsResult[2])

    return stsResult[0]
'''