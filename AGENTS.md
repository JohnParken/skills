# 项目 Skill 开发约定

本文件适用于整个仓库。以下相对路径以本文件所在的仓库根目录为基准。

## 必须读取的开发规范

- 每次新建、修改、导入、迁移或发布 Skill，先完整读取并遵循 [Skill 开发规范](skills/skill-development-standards/SKILL.md)，再进行写入。
- 修改 Skill 的脚本、参考资料、静态资源、依赖清单或测试，同样属于 Skill 维护；不要仅在修改 `SKILL.md` 时执行规范。
- 即使该 Skill 未出现在技能选择列表中，也必须按上面的真实路径读取。不要只依赖名称或 description 的自动匹配。
- 开发规范同时通过 `.agents/skills/skill-development-standards` 的相对链接提供项目内发现入口；维护唯一源文件，禁止复制出另一份规范。
- 普通业务任务只执行所用 Skill；不要因此改写 Skill 定义。保留已有修改，避免改动无关 Skill 或擅自更新已有打包文件。

## 完成前检查

- 按开发规范核对目录、元信息、资源、依赖和能力声明，并执行相关测试。新增或修改可执行逻辑时必须验证实际行为，不要仅检查语法。
- 从仓库根目录运行以下命令，将示例目标替换为本次实际维护的 Skill：

```console
python -I -S -B skills/skill-development-standards/scripts/check_skill.py skills/docx --strict
```

- 新建 Skill 应通过严格检查；维护已有 Skill 时修复本次引入的问题，报告范围外的既有例外及其原因。不要把失败、跳过或未执行的检查报告为通过。
- 修改开发规范检查器时，从仓库根目录运行：

```console
python -I -S -B -m unittest discover -s skills/skill-development-standards/scripts -p 'test_check_skill.py' -v
```

- 使用兼容的 Python 解释器；检查器仅依赖 Python 3.10 及以上版本的标准库。各业务 Skill 继续遵循自身运行环境和依赖约束，禁止将其他 Skill 的依赖引入标准库限定的 Skill。
- 人工检查自动检查器未覆盖的语义、功能及环境要求。最终说明改了什么、检查结果、尚未验证或保留的例外。
- 仅清理本次产生的临时文件、缓存和日志；禁止借整理规范之名删除用户文件或已有许可证。

## 维护规范本身

只有用户明确要求维护开发规范时才修改 `skills/skill-development-standards/` 中的规则或检查逻辑。不要为了让不合规的目标通过检查而削弱规范。
