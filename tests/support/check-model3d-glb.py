"""Independent GLB header/accessor/mesh/material/rotation checks; no product code."""
import sys,json,struct,math
from pathlib import Path
expected={'cube':(36,12,8,6),'pyramid':(18,6,5,5),'octahedron':(24,8,6,8)}
reports=[]
for filename in sys.argv[1:]:
 data=Path(filename).read_bytes();magic,version,total=struct.unpack_from('<III',data);assert (magic,version,total)==(0x46546c67,2,len(data))
 size,kind=struct.unpack_from('<II',data,12);assert size%4==0 and kind==0x4e4f534a
 doc=json.loads(data[20:20+size]);offset=20+size;bin_size,bin_kind=struct.unpack_from('<II',data,offset);assert bin_kind==0x004e4942 and bin_size%4==0 and offset+8+bin_size==len(data)
 binary=data[offset+8:];assert doc['asset']['version']=='2.0';assert not any(k in doc for k in ['animations','skins','textures','images','extensionsRequired','extensionsUsed']);assert len(doc['buffers'])==1 and 'uri' not in doc['buffers'][0];assert len(binary)-doc['buffers'][0]['byteLength'] in range(4)
 assert set(doc['extras']) == {'workbookModel3d'};model=next(iter(doc['extras'].values()));sample=model['sample'];vertices,faces,unique_count,planes=expected[sample];assert doc['scene']==0 and doc['scenes']==[{'nodes':[0]}] and len(doc['nodes'])==1 and len(doc['meshes'])==1 and len(doc['meshes'][0]['primitives'])==1
 primitive=doc['meshes'][0]['primitives'][0];assert primitive=={'attributes':{'POSITION':0,'NORMAL':1},'indices':2,'material':0,'mode':4};assert len(doc['bufferViews'])==3 and len(doc['accessors'])==3
 vectors=[]
 for index in [0,1]:
  a=doc['accessors'][index];assert a['componentType']==5126 and a['type']=='VEC3' and a['count']==vertices
  view=doc['bufferViews'][a['bufferView']];assert view['buffer']==0 and view['target']==34962 and view['byteOffset']%4==0 and view['byteLength']==vertices*12 and view['byteOffset']+view['byteLength']<=doc['buffers'][0]['byteLength']
  vectors.append([struct.unpack_from('<fff',binary,view['byteOffset']+i*12) for i in range(vertices)])
 positions,normals=vectors;assert all(math.isfinite(x) and abs(x)<=1 for p in positions for x in p);assert len(set(positions))==unique_count
 for axis in range(3):assert doc['accessors'][0]['min'][axis]==min(p[axis] for p in positions) and doc['accessors'][0]['max'][axis]==max(p[axis] for p in positions)
 a=doc['accessors'][2];view=doc['bufferViews'][a['bufferView']];assert a['componentType']==5123 and a['type']=='SCALAR' and a['count']==faces*3;assert view['target']==34963 and view['byteLength']==faces*6
 indices=struct.unpack_from('<'+'H'*(faces*3),binary,view['byteOffset']);assert list(indices)==list(range(vertices)) and max(indices)<vertices
 plane_set=set();edges={}
 for i in range(faces):
  points=[positions[j] for j in indices[i*3:i*3+3]];u=[points[1][k]-points[0][k] for k in range(3)];v=[points[2][k]-points[0][k] for k in range(3)];cross=(u[1]*v[2]-u[2]*v[1],u[2]*v[0]-u[0]*v[2],u[0]*v[1]-u[1]*v[0]);length=math.sqrt(sum(x*x for x in cross));assert length>0
  normal=tuple(x/length for x in cross);center=tuple(sum(p[k] for p in points)/3 for k in range(3));assert sum(normal[k]*center[k] for k in range(3))>0
  for j in indices[i*3:i*3+3]:assert sum((normals[j][k]-normal[k])**2 for k in range(3))<1e-12
  plane_set.add(tuple(round(n,6) for n in normal))
  for left,right in zip(points,points[1:]+points[:1]):
   edge=tuple(sorted([left,right]));edges.setdefault(edge,[]).append((left,right))
 assert len(plane_set)==planes
 for uses in edges.values():assert len(uses)==2 and uses[0]==uses[1][::-1]
 # Independent Euler-to-matrix and quaternion-to-matrix equality.
 x,y,z=[model['rotation'][k]*math.pi/180 for k in ['x','y','z']];cx,sx,cy,sy,cz,sz=math.cos(x),math.sin(x),math.cos(y),math.sin(y),math.cos(z),math.sin(z)
 wanted=[[cz*cy,cz*sy*sx-sz*cx,cz*sy*cx+sz*sx],[sz*cy,sz*sy*sx+cz*cx,sz*sy*cx-cz*sx],[-sy,cy*sx,cy*cx]]
 qx,qy,qz,qw=doc['nodes'][0]['rotation'];assert abs(qx*qx+qy*qy+qz*qz+qw*qw-1)<1e-12
 actual=[[1-2*(qy*qy+qz*qz),2*(qx*qy-qz*qw),2*(qx*qz+qy*qw)],[2*(qx*qy+qz*qw),1-2*(qx*qx+qz*qz),2*(qy*qz-qx*qw)],[2*(qx*qz-qy*qw),2*(qy*qz+qx*qw),1-2*(qx*qx+qy*qy)]]
 assert all(abs(actual[i][j]-wanted[i][j])<1e-12 for i in range(3) for j in range(3))
 material=doc['materials'][0];assert material['alphaMode']=='OPAQUE' and material['doubleSided'] is False;pbr=material['pbrMetallicRoughness'];assert pbr['metallicFactor']==0 and pbr['roughnessFactor']==1
 rgba=[int(model['color'][i:i+2],16)/255 for i in [0,2,4]];linear=[v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4 for v in rgba]+[1];assert all(abs(a-b)<1e-12 for a,b in zip(pbr['baseColorFactor'],linear))
 reports.append({'sample':sample,'vertices':vertices,'triangles':faces,'uniquePositions':unique_count,'outwardClosedMesh':True,'normalsMaterialQuaternion':True})
print(json.dumps({'passed':True,'assets':reports}))
