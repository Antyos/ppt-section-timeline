template = """
        <contextMenu idMso="{group_name}">
            <menu idMso="ObjectsGroupMenu">
                <button 
                    id="ObjectsAddToGroup_{group_name}"
                    insertAfterMso="ObjectsUngroup"
                    imageMso="ObjectsGroup"
                    label="Add to Group"
                    supertip="Add object to group while preserving animations."
                    showImage="true"
                    showLabel="true"
                    onAction="ObjectsAddToGroupCmd"
                    getEnabled="CanAddToGroup"
                    />
                <menuSeparator id="ObjectsAddToGroupSeparator_{group_name}" insertAfterMso="ObjectsUngroup" />
            </menu>
        </contextMenu>
"""

context_menus = [
    "ContextMenuPicture",
    "ContextMenuAudio",
    "ContextMenuVideo",
    "ContextMenuShape",
    "ContextMenuInk",
    "ContextMenuObjectsGroup",
    "ContextMenuShapeConnector",
    "ContextMenuShapeFreeform",
    "ContextMenuChartArea",
    "ContextMenuGraphicsCompatibility",
    "ContextMenuActiveXControl",
    "ContextMenuTableWhole",
    "ContextMenuSmartArtBackground",
    "ContextMenuOfficeWebExtension",
]

print("".join(template.format(group_name=cm) for cm in context_menus))
