def UIform(
    title,
    description,
    prefix_default,
    suffix_default,
    duplicate_options,
    default_index=0,
    button_text="Set Values",
):
    import WWP_uiUtils as ui
    labels = []
    values = []
    for opt in duplicate_options or []:
        if len(opt) >= 2:
            labels.append(opt[0])
            values.append(opt[1])
    return ui.uiUtils_duplicate_view_options(
        option_labels=labels,
        option_values=values,
        default_index=default_index,
        title=title,
        description=description,
        prefix_default=prefix_default,
        suffix_default=suffix_default,
        ok_text=button_text or "Set Values",
    )
