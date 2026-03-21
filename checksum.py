import struct

def compute_checksum(data, seed):

    
    total = 0
    block_count = 0x900000 >> 10  # each block is 0x400 bytes, so 0x900000/0x400 = 0x2400 blocks
    
    for b in range(block_count):
        acc_1 = 0
        acc_2 = 0
        base = b * 0x400
        for i in range(0, 0x400, 16):
            acc_1 += struct.unpack_from('<q', data, base + i)[0]
            acc_2 += struct.unpack_from('<q', data, base + i + 8)[0]
        # mask to 64-bit
        block_sum = (acc_1 + acc_2) & 0xFFFFFFFFFFFFFFFF
        total = (total + block_sum ^ seed) & 0xFFFFFFFFFFFFFFFF
    
    # fold 64-bit to 32-bit
    result = (total // 0xFFFFFFFF) + (total & 0xFFFFFFFF)
    result = result & 0xFFFFFFFF
    return result


def patch_checksum(data):
    seed = struct.unpack_from('<I', data, 0x900190)[0]  # file offset 0x900190
    checksum = compute_checksum(data[0x190:0x900190], seed)

    data= data[:0x900194] + struct.pack('<I', checksum) + data[0x900198:]
    print(f"Checksum patched: {checksum:08X}")
    return data, checksum