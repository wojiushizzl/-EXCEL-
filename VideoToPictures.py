import cv2
import os


video_path = r'D:\python_zzl\venv\EVA_PICTURE/'
cap = cv2.VideoCapture(video_path+'EVA.MP4')
frames_num=cap.get(7)
frames_fps=cap.get(5)
print(frames_num)
print(frames_fps)


def save_img():
    video_path = r'D:\python_zzl\venv\EVA_PICTURE/'
    videos = os.listdir(video_path)
    for video_name in videos:
        file_name = video_name.split('.')[0]
        folder_name = video_path + file_name
        os.makedirs(folder_name,exist_ok=True)
        vc = cv2.VideoCapture(video_path+video_name) #读入视频文件
        c=0
        rval=vc.isOpened()

        while rval:   #循环读取视频帧
            c = c + 1
            rval, frame = vc.read()
            pic_path = folder_name+'/'
            if rval:

                if c % 10 == 0:#原视频FPS/10提取图片
                    cv2.imwrite(pic_path + file_name + '_' + str(c/10) + '.jpg', frame) #存储为图像,保存名为 文件夹名_数字（第几个文件）.jpg
                    print(str(int(c/10))+' / '+str(int(frames_num/10)),  str(c/frames_num*100)+"%")
                    cv2.waitKey(1)
            else:
                break
        vc.release()
        print('save_success')
        print(folder_name)
save_img()