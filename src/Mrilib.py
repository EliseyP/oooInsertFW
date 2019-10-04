def Mri_test():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    document = XSCRIPTCONTEXT.getDocument()
    mri(ctx, document)


def mri(ctx, target):
    mri = ctx.ServiceManager.createInstanceWithContext(
        "mytools.Mri",ctx)
    mri.inspect(target)
