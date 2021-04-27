import math
from scipy.optimize import fsolve
from sympy import Symbol, nsolve, tan, sqrt
import numpy as np
from mayavi import mlab
import os
import math


def solve_equation(a_norm,b_norm,d_norm,er,eval_kz,eval_k0):
    kx = math.pi / a_norm
    ky = math.pi / b_norm
    kz = Symbol('kz')
    k0 = Symbol('k0')
    f1 = kz**2+kx**2+ky**2-er*(k0**2)
    f2 = kz*tan(kz*d_norm/2)-sqrt((er - 1)*(k0**2)-kz**2)
    #200 140
    #100 100
    #200 200
    root = nsolve((f1,f2),(kz,k0),(eval_kz,eval_k0))
    e_1 = float(root[0]**2+kx**2+ky**2-er*(root[1]**2))
    e_2 = float(root[0]*tan(root[0]*d_norm/2)-sqrt((er - 1)*(root[1]**2)-root[0]**2))
    if (np.isclose([e_1,e_2], [0.0, 0.0])).all():
        return root
    else:
        print("cannot find root",a_norm,b_norm,d_norm,er)
        return [0,0]

def plot_theoritial_field(x,y,z,kx,ky,kz,file_name,direction,save,even_mode):
    if even_mode:
        ex = ky*np.cos(kx * x)*np.cos(ky * y)*np.cos(kz * z)
        ey = kx*np.sin(kx * x)*np.sin(ky * y)*np.cos(kz * z)
        ez = np.zeros((21, 21, 21))
        if direction == 'y' or direction == 'z':
            ex = ky*np.sin(kx * x)*np.sin(ky * y)*np.cos(kz * z)
            ey = kx*np.cos(kx * x)*np.cos(ky * y)*np.cos(kz * z)
            ez = np.zeros((21, 21, 21))
    else:
        ex = ky*np.cos(kx * x)*np.sin(ky * y)*np.cos(kz * z)
        ey = -kx*np.sin(kx * x)*np.cos(ky * y)*np.cos(kz * z)
        ez = np.zeros((21, 21, 21))
    if direction == 'x':
        ez = ex
        ex = np.zeros((21, 21, 21))

    if direction == 'y':
        ez = ey
        ey = np.zeros((21, 21, 21))

    mag = np.sqrt(np.add(np.square(ex),np.square(ey),np.square(ez)))
    '''
    if save:
        v = np.stack((ex,ey,ez),axis=3)
        #v=np.concatenate((ex,ey,ez),axis=3)
        #print(v.shape)
        #print(ex)
        #print(ey)
        #print(ez)
        #print(v)
        #v = np.array([ex,ey,ez])
        #print(v.shape)
        np.save(file_name+"_info.npy",v)
    '''
    mlab.figure(bgcolor=(1,1,1))
    mlab.pipeline.vector_cut_plane(mlab.pipeline.vector_field(ex,ey,ez), mask_points=5, scale_factor=3, plane_orientation='x_axes')
    mlab.outline()
    mlab.savefig(filename=file_name+"_Vector_yz.png")
    mlab.show()
    mlab.pipeline.vector_cut_plane(mlab.pipeline.vector_field(ex,ey,ez), mask_points=5, scale_factor=3, plane_orientation='y_axes')
    mlab.outline()
    mlab.savefig(filename=file_name+"_Vector_xz.png")
    mlab.show()
    mlab.pipeline.vector_cut_plane(mlab.pipeline.vector_field(ex,ey,ez), mask_points=5, scale_factor=3, plane_orientation='z_axes')

    mlab.outline()
    mlab.savefig(filename=file_name+"_Vector__xy.png")
    mlab.show()

    '''
    mlab.pipeline.image_plane_widget(mlab.pipeline.scalar_field(mag),
                            plane_orientation='x_axes',
                            slice_index=11,
                        )

    #mlab.savefig(filename=file_name+"_yz.png")
    mlab.show()
    mlab.pipeline.image_plane_widget(mlab.pipeline.scalar_field(mag),
                            plane_orientation='y_axes',
                            slice_index=11,
                        )
    #mlab.savefig(filename=file_name+"_xz.png")
    mlab.show()
    mlab.pipeline.image_plane_widget(mlab.pipeline.scalar_field(mag),
                            plane_orientation='z_axes',
                            slice_index=11,
                        )
    #mlab.savefig(filename=file_name+"_xy.png")
    mlab.show()
    '''


def auto_plot(a,b,d,er,m,n,p,show,direction,save = False):
    '''
    a = 0.0319
    b = 0.0318
    d = 0.0261
    er = 10
    '''

    max_mode = max(m,n,p)
    if max_mode % 2 == 0:
        even_mode = True
    else:
        even_mode = False
    a_norm = a/m
    b_norm = b/n
    d_norm = d/p
    if direction == 'z' or direction == 'a':
        kx = math.pi / a_norm
        ky = math.pi / b_norm
        kz = math.pi / d_norm
        eval_k0 = math.sqrt((kx**2 + ky**2 + 0.75*kz**2)/er)

        kz,k0 = solve_equation(a_norm,b_norm,d_norm,er,kz * 0.8,eval_k0)
        kz = float(kz)
        print(kz)
        print(k0)
        f0 = float(k0)*3/20/math.pi
        f0 = round(f0,3)
        #print(f0)
        x, y, z = np.ogrid[-a/2:a/2:21j, -b/2:b/2:21j, -d/2:d/2:21j]

        file_name = "TE"+str(m)+str(n)+str(p)+"(z)"+'_'+str(f0)+'GHz'
        _base_path = os.getcwd()
        file_name = os.path.join(_base_path, 'Reference',file_name)
        print(file_name)
        if show:
            plot_theoritial_field(x,y,z,kx,ky,kz,file_name,'z',save,even_mode)
        if direction == 'z':
            return f0

    if direction == 'y' or direction == 'a':
        kx = math.pi / a_norm
        ky = math.pi / b_norm
        kz = math.pi / d_norm
        eval_k0 = math.sqrt((kx**2 + 0.75*ky**2 + kz**2)/er)
        ky,k0 = solve_equation(a_norm,d_norm,b_norm,er,ky * 0.8,eval_k0)
        ky = float(ky)
        f0 = float(k0)*3/20/math.pi
        f0 = round(f0,3)
        #print(f0)
        x, y, z = np.ogrid[-a/2:a/2:21j, -b/2:b/2:21j, -d/2:d/2:21j]
        kx = math.pi / a_norm
        kz = math.pi / d_norm

        file_name = "TE"+str(m)+str(n)+str(p)+"(y)"+'_'+str(f0)+'GHz'
        _base_path = os.getcwd()
        file_name = os.path.join(_base_path, 'Reference',file_name)
        print(file_name)
        if show:
            plot_theoritial_field(x,z,y,kx,kz,ky,file_name,'y',save,even_mode)
        if direction == 'y':
            return f0

    if direction == 'x' or direction == 'a':
        kx = math.pi / a_norm
        ky = math.pi / b_norm
        kz = math.pi / d_norm
        eval_k0 = math.sqrt((0.75*kx**2 + ky**2 + kz**2)/er)
        kx,k0 = solve_equation(d_norm,b_norm,a_norm,er,kx * 0.8,eval_k0)
        kx = float(kx)
        f0 = float(k0)*3/20/math.pi
        f0 = round(f0,3)
        #print(f0)
        x, y, z = np.ogrid[-a/2:a/2:21j, -b/2:b/2:21j, -d/2:d/2:21j]


        file_name = "TE"+str(m)+str(n)+str(p)+"(x)"+'_'+str(f0)+'GHz'
        _base_path = os.getcwd()
        file_name = os.path.join(_base_path, 'Reference',file_name)
        print(file_name)
        if show:
            plot_theoritial_field(z,y,x,kz,ky,kx,file_name,'x',save,even_mode)
        if direction == 'x':
            return f0


if __name__ == '__main__':
    a = 0.0142
    b = 0.012
    d = 0.0035
    er = 45

    m = 1
    n = 1
    p = 1
    auto_plot(a,b,d,er,m,n,p,True,'z',False)
