app' in the lower right of your app).
Traceback:
File "/mount/src/base-de-dados-pcp/app.py", line 239, in <module>
    df=previsao.merge(saldo,on="Codigo",how="left")\
       ~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/frame.py", line 12888, in merge
    return merge(
        self,
    ...<10 lines>...
        validate=validate,
    )
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/reshape/merge.py", line 385, in merge
    op = _MergeOperation(
        left_df,
    ...<10 lines>...
        validate=validate,
    )
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/reshape/merge.py", line 1018, in __init__
    ) = self._get_merge_keys()
        ~~~~~~~~~~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/reshape/merge.py", line 1633, in _get_merge_keys
    left_keys.append(left._get_label_or_level_values(lk))
                     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~^^^^
File "/home/adminuser/venv/lib/python3.14/site-packages/pandas/core/generic.py", line 1776, in _get_label_or_level_values
    raise KeyError(key)
