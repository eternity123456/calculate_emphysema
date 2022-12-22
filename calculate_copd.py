import SimpleITK as sitk
import os
import tqdm
import numpy as np
import collections
import xlwt
import copy


def get_listdir(path):
    tmp_list = []
    for file in os.listdir(path):
        if os.path.splitext(file)[1] == '.gz':
            file_path = os.path.join(path, file)
            tmp_list.append(file_path)
    return tmp_list


def volume(img, mask):
    img_sitk_img = sitk.ReadImage(img)
    img_arr = sitk.GetArrayFromImage(img_sitk_img)
    mask_sitk_img = sitk.ReadImage(mask)
    mask_arr = sitk.GetArrayFromImage(mask_sitk_img)
    Spacing = img_sitk_img.GetSpacing()
    voxel_volume = Spacing[0] * Spacing[1] * Spacing[2]

    lobe_count = collections.Counter(mask_arr.flatten())
    lobe1_volume = lobe_count[1] * voxel_volume
    lobe2_volume = lobe_count[2] * voxel_volume
    lobe3_volume = lobe_count[3] * voxel_volume
    lobe4_volume = lobe_count[4] * voxel_volume
    lobe5_volume = lobe_count[5] * voxel_volume

    copd_arr = np.zeros_like(img_arr)
    copd_arr[img_arr < -950] = 1

    lobe1 = copy.deepcopy(copd_arr)
    lobe1[mask_arr != 1] = 0
    lobe2 = copy.deepcopy(copd_arr)
    lobe2[mask_arr != 2] = 0
    lobe3 = copy.deepcopy(copd_arr)
    lobe3[mask_arr != 3] = 0
    lobe4 = copy.deepcopy(copd_arr)
    lobe4[mask_arr != 4] = 0
    lobe5 = copy.deepcopy(copd_arr)
    lobe5[mask_arr != 5] = 0

    lobe1_copd_volume = lobe1.sum() * voxel_volume
    lobe2_copd_volume = lobe2.sum() * voxel_volume
    lobe3_copd_volume = lobe3.sum() * voxel_volume
    lobe4_copd_volume = lobe4.sum() * voxel_volume
    lobe5_copd_volume = lobe5.sum() * voxel_volume

    return [lobe1_volume, lobe1_copd_volume / lobe1_volume,
            lobe2_volume, lobe2_copd_volume / lobe2_volume,
            lobe3_volume, lobe3_copd_volume / lobe3_volume,
            lobe1_volume + lobe2_volume + lobe3_volume,
            (lobe1_copd_volume + lobe2_copd_volume + lobe3_copd_volume) / (lobe1_volume + lobe2_volume + lobe3_volume),
            lobe4_volume, lobe4_copd_volume / lobe4_volume,
            lobe5_volume, lobe5_copd_volume / lobe5_volume,
            lobe4_volume + lobe5_volume,
            (lobe4_copd_volume + lobe5_copd_volume) / (lobe4_volume + lobe5_volume),
            lobe1_volume + lobe2_volume + lobe3_volume + lobe4_volume + lobe5_volume,
            (lobe1_copd_volume + lobe2_copd_volume + lobe3_copd_volume + lobe4_copd_volume + lobe5_copd_volume) / (
                    lobe1_volume + lobe2_volume + lobe3_volume + lobe4_volume + lobe5_volume)]


if __name__ == '__main__':
    img_path = r'G:\yxy\nii1'#修改为CT图像路径
    mask_path = r'G:\yxy\jieguo1'#修改为分割肺叶结果路径
    img_list = get_listdir(img_path)
    mask_list = get_listdir(mask_path)
    img_list.sort()
    mask_list.sort()
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('emphysema')
    value_name = ['path', 'RU volume', 'RU emphysema index', 'RM volume', 'RM emphysema index',
                  'RL volume', 'RL emphysema index', 'R volume', 'R emphysema index',
                  'LU volume', 'LU emphysema index', 'LL volume', 'LL emphysema index',
                  'L volume', 'L emphysema index', 'all lung volume', 'all lung em index']
    for value in range(len(value_name)):
        worksheet.write(0, value, value_name[value])
    for i in tqdm.trange(len(img_list)):
        hu_volume = volume(img_list[i], mask_list[i])
        worksheet.write(i + 1, 0, label=mask_list[i])
        for j in range(1, len(value_name)):
            worksheet.write(i + 1, j, label=hu_volume[j-1])

    workbook.save(r'emphysema_1.xls')#这里名字需要每次运行修改一下
