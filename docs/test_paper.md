# 封面

## 摘要

**摘要**：目标检测是计算机视觉领域的核心任务之一，广泛应用于自动驾驶、安防监控、工业检测等场景。近年来，基于深度学习的目标检测算法取得了显著进展，其中YOLO系列算法凭借其出色的速度与精度平衡成为研究热点。本文以YOLOv8为基础，针对小目标检测精度不足和复杂场景下漏检率较高的问题，提出了一种改进的实时目标检测算法。主要工作包括：在主干网络中引入注意力机制以增强特征提取能力；设计多尺度特征融合模块提升小目标检测性能；采用改进的损失函数优化训练过程。在COCO数据集上的实验结果表明，改进后的算法在保持实时检测速度的同时，$mAP$ 提升了 $3.2\%$，尤其在小目标检测任务上提升显著。

**关键词**：目标检测；YOLOv8；注意力机制；特征融合；深度学习

## 目录

# 第一章 绪论

## 一、研究背景与意义

目标检测是计算机视觉中最基础也最重要的任务之一，其目标是在给定图像中定位并识别出所有感兴趣的目标物体。随着深度学习技术的飞速发展，基于卷积神经网络的目标检测算法已经取得了突破性的进展。

在实际应用中，目标检测技术发挥着越来越重要的作用。在自动驾驶领域，车辆需要实时准确地识别道路上的行人、车辆和交通标志[1]；在安防监控领域，智能摄像头需要自动检测异常行为和可疑人员[2]；在工业生产中，自动化检测系统需要快速识别产品缺陷以保障质量[3]。这些应用场景对检测算法的速度和精度提出了极高的要求。

## 二、国内外研究现状

目标检测算法大致可以分为两类：基于候选区域的两阶段检测算法和基于回归的单阶段检测算法。

两阶段算法以R-CNN系列为代表。Girshick等人提出的R-CNN通过选择性搜索生成候选区域，再使用CNN提取特征并进行分类[4]。Fast R-CNN引入ROI Pooling实现了端到端训练[5]。Faster R-CNN进一步提出了区域建议网络（RPN），将候选区域生成也纳入网络中，大幅提升了检测速度[6]。

单阶段算法直接从图像中回归目标的类别和位置，代表性工作包括SSD[7]和YOLO系列。其中YOLO（You Only Look Once）算法由Redmon等人于2016年首次提出[8]，因其极高的检测速度而受到广泛关注。从YOLOv1到YOLOv8，该系列算法在网络结构、训练策略和后处理等方面不断改进，逐步提升了检测精度和速度。

## 三、本文主要工作

本文在YOLOv8的基础上进行改进，主要贡献包括：

1. 研究动机：分析YOLOv8在小目标检测上的不足，提出改进方案
2. 方法设计：引入注意力机制和改进的特征融合策略
3. 实验验证：在公开数据集上进行充分的实验对比与分析

> 注意：本文所有实验均在相同的硬件和软件环境下完成，以保证对比的公平性。

## 四、论文组织结构

本文共分为五章。第一章为绪论，介绍研究背景和国内外现状。第二章介绍相关理论基础。第三章详细描述改进算法的设计。第四章给出实验结果与分析。第五章总结全文并展望未来工作。

# 第二章 相关理论基础

## 一、卷积神经网络

卷积神经网络（Convolutional Neural Network，CNN）是深度学习中最成功的网络架构之一。一个典型的CNN由卷积层、池化层和全连接层组成。

卷积层通过可学习的滤波器对输入特征图进行卷积操作，提取局部特征。其计算过程可以表示为：

$$y(n) = \sum_{k} x(n+k) \cdot w(k) + b$$

其中 $x$ 为输入特征图，$w$ 为卷积核权重，$b$ 为偏置项。

池化层用于降低特征图的空间分辨率，减少计算量并增强特征的平移不变性。常用的池化操作包括最大池化和平均池化。

### （一）卷积运算

二维卷积的数学定义为：

$$y(i,j) = \sum_{m} \sum_{n} x(i+m, j+n) \cdot w(m,n) + b$$

在深度学习框架中，卷积运算还涉及步长（stride）和填充（padding）两个重要参数。

### （二）批量归一化

批量归一化（Batch Normalization）能够加速网络训练并提高稳定性：

$$\hat{x}_i = \frac{x_i - \mu_B}{\sqrt{\sigma_B^2 + \epsilon}}$$

$$y_i = \gamma \cdot \hat{x}_i + \beta$$

其中 $\mu_B$ 和 $\sigma_B^2$ 分别为小批量的均值和方差，$\gamma$ 和 $\beta$ 为可学习参数，$\epsilon$ 为防止除零的小常数。

## 二、目标检测评价标准

目标检测算法的性能通常通过以下指标进行评价：

| 指标名称 | 英文缩写 | 含义 |
| --- | --- | --- |
| 精确率 | Precision | 正确检测数占总检测数的比例 |
| 召回率 | Recall | 正确检测数占真实目标数的比例 |
| 平均精度均值 | mAP | 各类别AP值的均值 |
| 帧率 | FPS | 每秒处理的图像帧数 |
| 参数量 | Params | 模型中可训练参数的总数量 |

其中 $mAP$ 是最常用的综合评价指标，通常在 $IoU$ 阈值为 $0.5$ 时计算（$mAP@0.5$）。

精确率和召回率的计算公式分别为：

$$Precision = \frac{TP}{TP + FP}$$

$$Recall = \frac{TP}{TP + FN}$$

其中 $TP$ 为真正例，$FP$ 为假正例，$FN$ 为假负例。

## 三、YOLO系列算法概述

YOLO算法的核心思想是将目标检测任务转化为一个回归问题，直接从输入图像预测目标的类别概率和边界框坐标。相比两阶段算法，YOLO具有更快的检测速度，更适合实时应用场景。

YOLOv8是Ultralytics公司于2023年发布的最新版本，在网络结构上做了重大改进。其主干网络采用CSPDarknet结构，颈部网络使用PANet进行多尺度特征融合，检测头采用解耦设计，将分类和回归任务分离处理。

## 四、注意力机制

注意力机制能够让模型自适应地关注输入中的重要信息，抑制无关特征。常见的注意力机制包括：

- 通道注意力（SE-Net）：通过学习各通道的权重来增强有用特征
- 空间注意力（CBAM）：关注特征图中重要空间位置
- 自注意力（Transformer）：捕获全局依赖关系

SE模块的通道权重计算公式为：

$$s = \sigma(W_2 \cdot \delta(W_1 \cdot z))$$

其中 $z$ 为全局平均池化后的通道描述符，$W_1$ 和 $W_2$ 为全连接层权重，$\delta$ 为ReLU激活函数，$\sigma$ 为Sigmoid函数。

本文在YOLOv8的骨干网络中引入了一种轻量级的注意力模块，在不显著增加计算量的前提下提升特征表达能力。

# 第三章 改进的YOLOv8目标检测算法

## 一、算法整体框架

本文提出的改进算法在YOLOv8的基础上进行了三方面的优化，整体框架如下：

1. 主干网络改进：在C2f模块后插入注意力模块
2. 特征融合改进：增加小目标检测层
3. 损失函数改进：引入Focal-EIoU损失

## 二、注意力增强的主干网络

原始YOLOv8的主干网络采用CSPDarknet结构，虽然具有较好的特征提取能力，但在处理小目标和密集目标时仍存在不足。本文在主干网络的C2f模块之后引入了轻量级注意力模块（LAM），其结构包含以下步骤：

> 首先对输入特征图进行全局平均池化和最大池化，分别得到通道描述符；然后将两者拼接后经过共享的全连接层，生成通道注意力权重；最后将权重作用于原始特征图。

该模块的计算开销仅增加约 $0.3\%$ 的参数量，但能够有效增强网络对关键特征的关注能力。

LAM模块的输出可以表示为：

$$\mathbf{F}' = \mathbf{F} \otimes \text{LAM}(\mathbf{F})$$

其中 $\mathbf{F}$ 为输入特征图，$\otimes$ 表示逐元素乘法，$\text{LAM}(\cdot)$ 为注意力权重生成函数。

## 三、多尺度特征融合改进

原始YOLOv8使用三个尺度的特征图进行检测，分别对应大、中、小三种尺寸的目标。然而，对于极小目标（像素面积小于 $32 \times 32$），仅依靠这三个尺度往往难以提取到足够的细节信息。

本文在PANet的基础上增加了一个第四尺度的检测头，专门用于检测极小目标。具体做法是将backbone中更浅层的特征图（分辨率更大）与深层特征图进行融合，从而保留更多的小目标细节信息。

**表 3-1 改进前后检测尺度对比**

| 检测尺度 | 原始YOLOv8 | 改进算法 | 下采样倍率 |
| --- | --- | --- | --- |
| 第一尺度 | P5 | P5 | 32× |
| 第二尺度 | P4 | P4 | 16× |
| 第三尺度 | P3 | P3 | 8× |
| 第四尺度 | 无 | P2 | 4× |

## 四、损失函数设计

原始YOLOv8使用CIoU Loss作为边界框回归损失。本文将其替换为Focal-EIoU Loss，该损失函数在EIoU的基础上引入了Focal机制，能够使模型更加关注难检测的样本。

Focal-EIoU Loss的计算公式如下：

$$L_{\text{focal-eiou}} = \text{IoU}^{\gamma} \cdot (1 - \text{IoU}) \cdot L_{\text{EIoU}}$$

其中 $IoU$ 为预测框与真实框的交并比，$\gamma$ 为聚焦参数（本文取 $2$），$L_{\text{EIoU}}$ 包含宽高比和中心点距离的信息。

EIoU的完整计算公式为：

$$
L_{\text{EIoU}} = 1 - \text{IoU} + \frac{\rho^2(\mathbf{b}, \mathbf{b}^{gt})}{c^2} + \frac{\rho^2(w, w^{gt})}{C_w^2} + \frac{\rho^2(h, h^{gt})}{C_h^2}
$$

其中 $\rho(\cdot)$ 为欧氏距离，$\mathbf{b}$ 和 $\mathbf{b}^{gt}$ 分别为预测框和真实框的中心点，$c$ 为两框最小包围矩形对角线长度，$C_w$ 和 $C_h$ 为包围矩形的宽和高。

## 五、训练策略

本文采用以下训练策略以提升模型性能：

1. 数据增强：使用Mosaic、MixUp、随机裁剪和颜色抖动
2. 学习率调度：余弦退火策略，初始学习率为 $0.01$
3. 预训练权重：使用在COCO上预训练的YOLOv8权重作为初始化
4. 训练轮次：共训练 $300$ 个epoch，batch size为 $16$

# 第四章 实验结果与分析

## 一、实验环境与数据集

本文实验在以下环境中进行：

| 项目 | 配置 |
| --- | --- |
| 操作系统 | Ubuntu 22.04 LTS |
| GPU | NVIDIA RTX 4090 24GB |
| CPU | Intel i9-13900K |
| 内存 | 64GB DDR5 |
| 深度学习框架 | PyTorch 2.1.0 |
| CUDA版本 | 12.1 |

实验使用MS COCO 2017数据集进行训练和测试，该数据集包含约12万张训练图像和5000张验证图像，涵盖80个目标类别。

## 二、消融实验

为验证各改进模块的有效性，本文设计了消融实验，结果如下：

**表 4-1 消融实验结果**

| 模型 | 注意力模块 | P2检测层 | Focal-EIoU | $mAP@0.5$ | FPS |
| --- | --- | --- | --- | --- | --- |
| YOLOv8 baseline | - | - | - | 53.9% | 127 |
| + 注意力模块 | ✓ | - | - | 54.8% | 124 |
| + P2检测层 | - | ✓ | - | 55.3% | 112 |
| + Focal-EIoU | - | - | ✓ | 55.1% | 127 |
| 全部改进（本文方法） | ✓ | ✓ | ✓ | 57.1% | 109 |

从表4-1可以看出，每个改进模块都能带来一定的性能提升，三者结合后 $mAP@0.5$ 达到 $57.1\%$，相比基线提升了 $3.2\%$，同时保持了较好的实时性能（109 FPS）。

改进算法的总体性能提升可以量化为：

$$\Delta mAP = mAP_{\text{ours}} - mAP_{\text{baseline}} = 57.1\% - 53.9\% = 3.2\%$$

## 三、与主流算法对比

将本文算法与其他主流目标检测算法进行对比：

**表 4-2 与主流算法对比结果**

| 算法 | 骨干网络 | $mAP@0.5$ | FPS | 参数量(M) |
| --- | --- | --- | --- | --- |
| Faster R-CNN | ResNet-50 | 50.2% | 18 | 41.1 |
| SSD300 | VGG-16 | 47.5% | 46 | 26.3 |
| YOLOv5s | CSPDarknet | 50.7% | 140 | 7.2 |
| YOLOv8s | CSPDarknet | 53.9% | 127 | 11.2 |
| DETA-Net | ResNet-50 | 54.6% | 35 | 39.8 |
| 本文方法 | CSPDarknet+LAM | 57.1% | 109 | 12.8 |

## 四、小目标检测分析

为了验证改进算法在小目标检测上的效果，本文在COCO数据集上按照目标大小分别统计了检测结果。

各尺寸目标的检测精度如下：

**表 4-3 不同尺寸目标检测AP对比**

| 目标尺寸 | 基线AP | 改进AP | 提升 |
| --- | --- | --- | --- |
| 小型（$area < 32^2$） | 28.3% | 34.1% | +5.8% |
| 中型（$32^2 < area < 96^2$） | 51.7% | 53.8% | +2.1% |
| 大型（$area > 96^2$） | 63.5% | 65.0% | +1.5% |

在小型目标上，基线算法的AP为 $28.3\%$，本文方法达到 $34.1\%$，提升了 $5.8$ 个百分点。这说明增加P2检测层和注意力机制对小目标检测的提升最为显著。

## 五、可视化分析

通过对比检测结果的可视化效果可以发现，改进后的算法在以下场景中表现更优：

- 远距离小目标：能够检测到基线算法遗漏的远处行人和车辆
- 密集目标：在拥挤场景中减少了漏检
- 遮挡目标：对部分遮挡的目标具有更强的鲁棒性

# 第五章 总结与展望

## 一、本文工作总结

本文针对YOLOv8在小目标检测和复杂场景下的不足，提出了三点改进：在主干网络中引入轻量级注意力模块增强特征表达；增加P2检测层提升小目标检测能力；采用Focal-EIoU损失函数优化训练过程。实验结果表明改进算法在COCO数据集上取得了显著提升。

总体改进效果可以概括为：

$$\text{Performance} = f(\text{Attention}, \text{MultiScale}, \text{Loss})$$

## 二、不足与展望

尽管取得了一定成果，本文仍存在以下不足和可改进之处：

1. 在极端密集场景下仍存在漏检现象
2. 模型参数量和计算量有所增加
3. 仅在COCO数据集上验证，泛化能力有待进一步考察

未来工作将从以下方向展开：探索更高效的网络结构以减少计算开销；将改进算法应用于实际场景（如自动驾驶）进行验证；研究基于Transformer的检测头以进一步提升性能。

## 参考文献

[1] Li J, Zhang Y, Wang S. Deep learning for autonomous driving: A survey. IEEE Transactions on Intelligent Vehicles, 2023, 8(4): 2982-3000.

[2] Zhang W, Chen L, Liu H. Intelligent video surveillance: A review of deep learning-based approaches. Pattern Recognition, 2022, 130: 108781.

[3] Wang R, Huang J, Chen X. Industrial defect detection based on improved YOLOv5. Journal of Manufacturing Systems, 2023, 66: 234-245.

[4] Girshick R, Donahue J, Darrell T, et al. Rich feature hierarchies for accurate object detection and semantic segmentation. In: CVPR. 2014: 580-587.

[5] Girshick R. Fast R-CNN. In: ICCV. 2015: 1440-1448.

[6] Ren S, He K, Girshick R, et al. Faster R-CNN: Towards real-time object detection with region proposal networks. IEEE TPAMI, 2017, 39(6): 1137-1149.

[7] Liu W, Anguelov D, Erhan D, et al. SSD: Single shot multibox detector. In: ECCV. 2016: 21-37.

[8] Redmon J, Divvala S, Girshick R, et al. You only look once: Unified, real-time object detection. In: CVPR. 2016: 779-788.

[9] Redmon J, Farhadi A. YOLO9000: Better, faster, stronger. In: CVPR. 2017: 7263-7271.

[10] Redmon J, Farhadi A. YOLOv3: An incremental improvement. arXiv preprint arXiv:1804.02767, 2018.

[11] Bochkovskiy A, Wang C Y, Liao H Y M. YOLOv4: Optimal speed and accuracy of object detection. arXiv preprint arXiv:2004.10934, 2020.

[12] Jocher G, Chaurasia A, Qiu J. Ultralytics YOLOv8. 2023.

[13] Hu J, Shen L, Sun G. Squeeze-and-excitation networks. In: CVPR. 2018: 7132-7141.

[14] Woo S, Park J, Lee J Y, et al. CBAM: Convolutional block attention module. In: ECCV. 2018: 3-19.

[15] Lin T Y, Dollár P, Girshick R, et al. Feature pyramid networks for object detection. In: CVPR. 2017: 2117-2125.
